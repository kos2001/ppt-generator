#!/usr/bin/env python
"""Resize pictures in a PPTX by automating desktop PowerPoint via COM (pywin32).

Unlike edit_pptx.py / python-pptx (which parse the file directly and CANNOT open
DRM-protected files), this drives the installed PowerPoint application. If the
current Windows user can open and edit the file manually in PowerPoint -
including files protected by a corporate DRM/EDM client - then PowerPoint
decrypts it in that authorized session, this script resizes the pictures, and
Save() lets the DRM agent re-encrypt on write. It does NOT bypass DRM; it relies
on PowerPoint's own authorized access with the user's existing edit rights.

Resizing changes only a picture's frame geometry (size/position) - it never
reads, decodes, or exports the image's pixel bytes. So even on decks where DRM
blocks image *extraction*, resizing still works (manual resize working is the
proof). If the DRM policy blocks automation/COM access entirely, the Open() call
fails - that block is an intended control; do not work around it.

Requirements:
    - Desktop Microsoft PowerPoint installed (COM-registered)
    - pip install pywin32
    - For DRM/EDM files: the corporate DRM client installed, user signed in,
      and edit+save rights on the document

Run it on the corporate PC (where PowerPoint + DRM client + the file live).

Workflow (do these in order):
    # 1) Can COM even open it under the DRM policy? (opens read-only, no edit)
    python scripts/resize_pptx_com.py deck.pptx --check

    # 2) See every picture and its current size, to decide targets/sizes
    python scripts/resize_pptx_com.py deck.pptx --list

    # 3) Resize. Examples:
    #    fit EVERY picture into the template's content box (house default):
    python scripts/resize_pptx_com.py deck.pptx --template
    #    one picture into the samsung template box, on slide 3:
    python scripts/resize_pptx_com.py deck.pptx --slide 3 --picture 1 --template samsung
    #    one picture to an exact width (height follows, aspect kept):
    python scripts/resize_pptx_com.py deck.pptx --slide 3 --picture 1 --width 10
    #    scale every picture to 80%:
    python scripts/resize_pptx_com.py deck.pptx --scale 0.8
    #    fit a picture inside a custom 20x12 cm box, centered:
    python scripts/resize_pptx_com.py deck.pptx --slide 2 --picture 1 --box 20x12
    #    save to a new file instead of editing in place:
    python scripts/resize_pptx_com.py deck.pptx --template --out resized.pptx

Sizes are in centimeters by default (--unit cm|in|pt). In-place edits back up
the original to <name>.bak-<timestamp>.pptx first.
"""
import argparse
import os
import shutil
import sys
from datetime import datetime

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")

try:
    import win32com.client
    from pywintypes import com_error
except ImportError:
    print("pywin32 is required: pip install pywin32", file=sys.stderr)
    sys.exit(1)

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from templates import body_box, DEFAULT_THEME  # shared template content box

# COM / Office constants
MSO_PICTURE = 13          # msoPicture
MSO_LINKED_PICTURE = 11   # msoLinkedPicture
MSO_TRUE = -1
MSO_FALSE = 0
PP_SAVE_PPTX = 24         # ppSaveAsOpenXMLPresentation

# Unit conversions. PowerPoint measures shapes in points (1 in = 72 pt).
UNIT_TO_PT = {"cm": 28.3464567, "in": 72.0, "pt": 1.0}
PT_TO_CM = 1.0 / 28.3464567
PT_TO_IN = 1.0 / 72.0


def _is_picture(shape):
    try:
        return shape.Type in (MSO_PICTURE, MSO_LINKED_PICTURE)
    except com_error:
        return False


def _pictures(slide):
    """Pictures on a slide, in z-order, as a list (1-based index = position)."""
    return [sh for sh in slide.Shapes if _is_picture(sh)]


def _open(path, *, read_only):
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    pres = powerpoint.Presentations.Open(
        os.path.abspath(path),
        ReadOnly=MSO_TRUE if read_only else MSO_FALSE,
        Untitled=MSO_FALSE,
        WithWindow=MSO_FALSE,
    )
    return powerpoint, pres


def _slide_dims_in(pres):
    """Slide width, height in inches (from PowerPoint's point-based PageSetup)."""
    ps = pres.PageSetup
    return ps.SlideWidth * PT_TO_IN, ps.SlideHeight * PT_TO_IN


def cmd_check(path):
    """Step 1: prove the DRM policy lets COM open the file at all."""
    try:
        powerpoint, pres = _open(path, read_only=True)
    except com_error as e:
        print("COM OPEN FAILED.")
        print("  ->", e)
        print("\nLikely causes:")
        print("  - DRM/EDM policy blocks automation (COM) access to this file")
        print("  - You lack rights to open it, or the DRM client isn't signed in")
        print("  - PowerPoint isn't installed / COM-registered")
        print("\nIf manual editing works but this fails, the policy blocks")
        print("automation on purpose - that's an intended control, not a bug.")
        return 1
    n_slides = pres.Slides.Count
    n_pics = sum(len(_pictures(pres.Slides(i))) for i in range(1, n_slides + 1))
    try:
        protected = bool(pres.Permission.Enabled)  # IRM/RMS permission set?
    except com_error:
        protected = None
    pres.Close()
    powerpoint.Quit()
    print("COM OPEN OK - PowerPoint opened the file in an authorized session.")
    print("  slides: %d   pictures: %d" % (n_slides, n_pics))
    if protected is True:
        print("  IRM/RMS permission: ENABLED (rights-managed document)")
    print("\nCOM access works. You can proceed to --list, then resize.")
    print("Note: --check opens read-only; saving still needs edit+save rights.")
    return 0


def cmd_list(path):
    """Step 2: inventory every picture and its current size (cm)."""
    powerpoint, pres = _open(path, read_only=True)
    try:
        for i in range(1, pres.Slides.Count + 1):
            pics = _pictures(pres.Slides(i))
            if not pics:
                continue
            print("Slide %d:" % i)
            for j, sh in enumerate(pics, start=1):
                print("  picture %d: %.2f x %.2f cm  @ (%.2f, %.2f cm)  name=%r"
                      % (j, sh.Width * PT_TO_CM, sh.Height * PT_TO_CM,
                         sh.Left * PT_TO_CM, sh.Top * PT_TO_CM, sh.Name))
    finally:
        pres.Close()
        powerpoint.Quit()
    return 0


def _fit_into_box(sh, box_l, box_t, box_w, box_h, mode):
    """Resize/reposition a picture to a fixed box (points). No pixel access.

    contain: scale to fit inside the box, centered, keeping aspect (no crop).
    stretch: fill the box exactly, distorting to match.
    Aspect is taken from the picture's current displayed size.
    """
    if mode == "stretch":
        sh.LockAspectRatio = MSO_FALSE
        sh.Left, sh.Top, sh.Width, sh.Height = box_l, box_t, box_w, box_h
        return
    cur_w, cur_h = sh.Width, sh.Height
    if not cur_w or not cur_h:
        return
    scale = min(box_w / cur_w, box_h / cur_h)
    new_w, new_h = cur_w * scale, cur_h * scale
    sh.LockAspectRatio = MSO_FALSE
    sh.Width, sh.Height = new_w, new_h
    sh.Left = box_l + (box_w - new_w) / 2.0
    sh.Top = box_t + (box_h - new_h) / 2.0


def _resize_freeform(sh, args, to_pt):
    """Scale / exact-size a picture (no fixed box)."""
    keep_aspect = not args.no_keep_aspect
    sh.LockAspectRatio = MSO_TRUE if keep_aspect else MSO_FALSE
    if args.scale is not None:
        # With aspect locked, setting Width auto-scales Height.
        sh.Width = sh.Width * args.scale
        if not keep_aspect:
            sh.Height = sh.Height * args.scale
    else:  # exact width and/or height
        if args.width is not None:
            sh.Width = args.width * to_pt
        if args.height is not None and (not keep_aspect or args.width is None):
            sh.Height = args.height * to_pt
    if args.left is not None:
        sh.Left = args.left * to_pt
    if args.top is not None:
        sh.Top = args.top * to_pt


def cmd_resize(path, args):
    to_pt = UNIT_TO_PT[args.unit]
    powerpoint, pres = _open(path, read_only=False)
    changed = 0
    try:
        # Resolve a fixed box (in points) for template / box modes, if any.
        box_pt = None
        if args.template is not None:
            sw, sh_in = _slide_dims_in(pres)
            l, t, w, h = body_box(args.template, sw, sh_in)
            box_pt = (l * 72, t * 72, w * 72, h * 72)
            print("Fitting into %s template box: %.2f x %.2f cm @ (%.2f, %.2f cm), "
                  "fit=%s" % (args.template, w * 2.54, h * 2.54,
                              l * 2.54, t * 2.54, args.fit))
        elif args.box is not None:
            bw, bh = (v * to_pt for v in args.box)
            bl = (args.left * to_pt) if args.left is not None else None
            bt = (args.top * to_pt) if args.top is not None else None
            box_pt = (bl, bt, bw, bh)  # bl/bt None -> per-picture current pos

        slide_range = ([args.slide] if args.slide
                       else range(1, pres.Slides.Count + 1))
        for i in slide_range:
            slide = pres.Slides(i)
            pics = _pictures(slide)
            if args.picture:
                if args.picture > len(pics):
                    print("  slide %d: no picture #%d (has %d)"
                          % (i, args.picture, len(pics)))
                    continue
                targets = [(args.picture, pics[args.picture - 1])]
            else:
                targets = list(enumerate(pics, start=1))
            for j, sh in targets:
                before = (sh.Width * PT_TO_CM, sh.Height * PT_TO_CM)
                if box_pt is not None:
                    bl, bt, bw, bh = box_pt
                    if bl is None:
                        bl = sh.Left  # anchor custom box at picture's position
                    if bt is None:
                        bt = sh.Top
                    _fit_into_box(sh, bl, bt, bw, bh, args.fit)
                else:
                    _resize_freeform(sh, args, to_pt)
                after = (sh.Width * PT_TO_CM, sh.Height * PT_TO_CM)
                print("  slide %d picture %d: %.2fx%.2f -> %.2fx%.2f cm"
                      % (i, j, before[0], before[1], after[0], after[1]))
                changed += 1

        if changed == 0:
            print("No pictures matched; nothing saved.")
            return 1

        out = args.out
        if out and os.path.abspath(out) != os.path.abspath(path):
            pres.SaveAs(os.path.abspath(out), PP_SAVE_PPTX)
            print("Saved -> %s (%d picture(s) resized)" % (out, changed))
        else:
            stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
            backup = "%s.bak-%s.pptx" % (os.path.splitext(path)[0], stamp)
            shutil.copy2(path, backup)
            print("Backed up original -> %s" % backup)
            pres.Save()
            print("Saved in place -> %s (%d picture(s) resized)" % (path, changed))
    except com_error as e:
        print("COM error during resize/save:", e, file=sys.stderr)
        print("If Open worked but Save failed, the DRM policy likely grants "
              "view/automation but not save rights.", file=sys.stderr)
        return 1
    finally:
        pres.Close()
        powerpoint.Quit()
    return 0


def main():
    ap = argparse.ArgumentParser(
        description="Resize pictures in a (possibly DRM-protected) PPTX via "
                    "PowerPoint COM automation.")
    ap.add_argument("path", help="path to .pptx")
    ap.add_argument("--check", action="store_true",
                    help="step 1: test that COM can open the file (read-only)")
    ap.add_argument("--list", action="store_true", dest="do_list",
                    help="step 2: list pictures and current sizes")
    ap.add_argument("--slide", type=int, help="target slide (1-based); default all")
    ap.add_argument("--picture", type=int,
                    help="target picture # on the slide (1-based); default all")
    # sizing modes (pick one)
    ap.add_argument("--template", nargs="?", const=DEFAULT_THEME, default=None,
                    metavar="NAME",
                    help="fit picture(s) into the template's content box "
                         "(no value = %s)" % DEFAULT_THEME)
    ap.add_argument("--box", help="fit into a custom WxH box, e.g. 20x12")
    ap.add_argument("--scale", type=float, help="multiply current size by FACTOR")
    ap.add_argument("--width", type=float, help="set exact width")
    ap.add_argument("--height", type=float, help="set exact height")
    # modifiers
    ap.add_argument("--fit", choices=("contain", "stretch"), default="contain",
                    help="how to fit into --template/--box (default contain)")
    ap.add_argument("--unit", choices=("cm", "in", "pt"), default="cm",
                    help="unit for sizes/positions (default cm)")
    ap.add_argument("--left", type=float, help="set left position")
    ap.add_argument("--top", type=float, help="set top position")
    ap.add_argument("--no-keep-aspect", action="store_true",
                    help="allow distortion in --scale/--width/--height modes")
    ap.add_argument("--out", help="save to a new file instead of in place")
    args = ap.parse_args()

    if args.box:
        try:
            w, h = (float(x) for x in args.box.lower().split("x"))
            args.box = (w, h)
        except ValueError:
            ap.error("--box must look like WxH, e.g. 20x12")

    modes = [args.template is not None, args.box is not None,
             args.scale is not None, args.width is not None or args.height is not None]
    if sum(bool(m) for m in modes) > 1:
        ap.error("choose ONE resize mode: --template, --box, --scale, or "
                 "--width/--height")

    try:
        if args.check:
            return cmd_check(args.path)
        if args.do_list:
            return cmd_list(args.path)
        if not any(modes):
            ap.error("choose a resize mode: --template, --box, --scale, or "
                     "--width/--height (or use --check / --list)")
        return cmd_resize(args.path, args)
    except com_error as e:
        print("Failed to open via PowerPoint:", e, file=sys.stderr)
        print("Ensure desktop PowerPoint is installed and (for DRM files) that "
              "you are signed into the DRM client with rights to this document.",
              file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
