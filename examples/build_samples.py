"""Build the reference examples with native Microsoft PowerPoint on Windows.

This example builder uses already supplied, provenance-tracked backgrounds.
It does not generate or search for new images, install fonts, or benchmark a model.
"""

import argparse
import json
import re
import time
from pathlib import Path

from PIL import Image
import pythoncom
import win32com.client


EXAMPLES = Path(__file__).resolve().parent
W, H = 960, 540
TITLE_FONT = "方正小标宋简体"
BODY_FONT = "仿宋"
LATIN_FONT = "Times New Roman"


def rgb(color):
    color = color.lstrip('#')
    red, green, blue = (int(color[i:i + 2], 16) for i in (0, 2, 4))
    return red + (green << 8) + (blue << 16)


def style_text(text_range, family, size, color):
    text_range.Font.Name = family
    text_range.Font.NameFarEast = family
    text_range.Font.NameAscii = LATIN_FONT
    text_range.Font.NameOther = LATIN_FONT
    text_range.Font.Size = size
    text_range.Font.Bold = 0
    text_range.Font.Color.RGB = rgb(color)
    # Explicit ranges preserve Latin digits/words across application font mappings.
    for match in re.finditer(r'[A-Za-z0-9]+(?:[ .:/%–-][A-Za-z0-9]+)*', text_range.Text):
        part = text_range.Characters(match.start() + 1, len(match.group()))
        part.Font.Name = LATIN_FONT
        part.Font.NameAscii = LATIN_FONT
        part.Font.Bold = 0


def textbox(slide, text, box, family, size, color, name):
    shape = slide.Shapes.AddTextbox(1, *box)
    shape.Name = name
    shape.TextFrame.AutoSize = 0
    shape.TextFrame.MarginLeft = 0
    shape.TextFrame.MarginRight = 0
    shape.TextFrame.MarginTop = 0
    shape.TextFrame.MarginBottom = 0
    shape.TextFrame.WordWrap = -1
    shape.TextFrame.AutoSize = 0
    shape.TextFrame.TextRange.Text = text
    style_text(shape.TextFrame.TextRange, family, size, color)
    # Setting margins can resize an empty PowerPoint textbox before text arrives.
    shape.Left, shape.Top, shape.Width, shape.Height = box
    return shape


def background(slide, filename, palette, strength):
    with Image.open(filename) as image:
        image_w, image_h = image.size
    ratio = max(W / image_w, H / image_h)
    picture = slide.Shapes.AddPicture(str(filename), 0, -1,
                                     (W - image_w * ratio) / 2,
                                     (H - image_h * ratio) / 2,
                                     image_w * ratio, image_h * ratio)
    picture.Name = 'background-image'
    picture.AlternativeText = palette['background_description']
    mask = slide.Shapes.AddShape(1, 0, 0, W, H)
    mask.Name = 'background-text-fade'
    mask.Line.Visible = 0
    # Office's vertical gradient style uses vertical bands, fading left to right.
    mask.Fill.TwoColorGradient(2, 1)
    for index, transparency in [(1, strength), (2, 1.0)]:
        stop = mask.Fill.GradientStops.Item(index)
        stop.Color.RGB = rgb(palette['paper'])
        stop.Transparency = transparency
    # A soft plateau keeps busy photography readable below the text column.
    mask.Fill.GradientStops.Insert(rgb(palette['paper']), 0.60, strength + 0.03)
    mask.AlternativeText = ''
    return picture


def animate(slide, shape, first=False):
    # MsoAnimEffectFade=10; with previous=2; on page click=1.
    effect = slide.TimeLine.MainSequence.AddEffect(shape, 10, 0, 2 if first else 1)
    effect.Timing.Duration = 0.45
    effect.Timing.TriggerDelayTime = 0.10 if first else 0.0


def table(slide, data, palette):
    rows = data['table']
    shape = slide.Shapes.AddTable(len(rows), len(rows[0]), 64, 170, 700, 240)
    shape.Name = 'editable-schedule'
    column_widths = data.get('column_widths', [145, 255, 300])
    for j, width in enumerate(column_widths, 1):
        shape.Table.Columns.Item(j).Width = width
    for i, row in enumerate(rows, 1):
        shape.Table.Rows.Item(i).Height = 60
        for j, value in enumerate(row, 1):
            cell = shape.Table.Cell(i, j)
            cell.Shape.Fill.Solid()
            cell.Shape.Fill.ForeColor.RGB = rgb(palette['paper'])
            cell.Shape.Fill.Transparency = 1.0
            for border in range(1, 5):
                cell.Borders.Item(border).ForeColor.RGB = rgb(palette['paper'])
                cell.Borders.Item(border).Transparency = 0.78
                cell.Borders.Item(border).Weight = 0.5
            frame = cell.Shape.TextFrame
            frame.MarginLeft, frame.MarginRight = 12, 10
            frame.MarginTop, frame.MarginBottom = 8, 8
            frame.VerticalAnchor = 3
            frame.TextRange.Text = value
            style_text(frame.TextRange, BODY_FONT, 20 if i == 1 else 21, palette['ink'])
    return shape


def build(variant, output_root, make_video):
    source = EXAMPLES / variant
    config = json.loads((source / 'slides.json').read_text(encoding='utf-8'))
    assets = json.loads((source / 'asset-manifest.json').read_text(encoding='utf-8'))['assets']
    dest = output_root / variant
    dest.mkdir(parents=True, exist_ok=True)
    deck_path = dest / 'sample.pptx'
    if deck_path.exists():
        raise FileExistsError(f'Choose a new output directory; output exists: {deck_path}')
    app = win32com.client.DispatchEx('PowerPoint.Application')
    pres = app.Presentations.Add(False)
    manifest = {'reference_builder': 'Native Microsoft PowerPoint COM',
                'powerpoint_version': str(app.Version), 'powerpoint_build': str(app.Build),
                'slides': []}
    try:
        pres.PageSetup.SlideWidth, pres.PageSetup.SlideHeight = W, H
        for index, data in enumerate(config['slides'], 1):
            slide = pres.Slides.Add(index, 12)
            palette = config['palette']
            asset = assets[index - 1]
            background(slide, source / asset['file'], palette,
                       config['fade_strength'] if data['kind'] != 'table' else 0.02)
            title_box = (64, 151, 650, 80) if data['kind'] == 'cover' else (64, 66, 800, 68)
            title = textbox(slide, data['title'], title_box, TITLE_FONT,
                            44 if data['kind'] == 'cover' else 34, palette['ink'], 'slide-title')
            animated = [title]
            if data['kind'] == 'cover':
                for j, line in enumerate(data['body']):
                    animated.append(textbox(slide, line, (66, 250 + j * 48, 590, 65),
                                            BODY_FONT, 24, palette['ink'], f'body-{j + 1}'))
            elif data['kind'] == 'table':
                animated.append(table(slide, data, palette))
                if data.get('caption'):
                    animated.append(textbox(slide, data['caption'], (66, 437, 740, 42),
                                            BODY_FONT, 20, palette['ink'], 'table-caption'))
            else:
                for j, line in enumerate(data['body']):
                    animated.append(textbox(slide, line, (66, 169 + j * 88, 620, 70),
                                            BODY_FONT, 24, palette['ink'], f'body-{j + 1}'))
            for j, shape in enumerate(animated):
                animate(slide, shape, j == 0)
            footer = textbox(slide, f'{index:02}', (66, 492, 45, 22), LATIN_FONT, 13,
                             palette['ink'], 'page-number')
            note_lines = [data.get('notes', ''), '正文来源：本示例 input.md 的原创内容。']
            if asset.get('source_page'):
                note_lines.append('图片出处：' + asset['source_page'])
                note_lines.append('作者：' + asset['author'] + '；许可：' + asset['license'])
            else:
                note_lines.append('图片资源：' + asset['file'])
            note_text = '\n'.join(note_lines)
            notes = slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange
            notes.Text = note_text
            style_text(notes, BODY_FONT, 12, palette['ink'])
            transition = slide.SlideShowTransition
            transition.EntryEffect = 1793
            transition.Duration = 0.45
            transition.AdvanceOnClick = -1
            transition.AdvanceOnTime = 0
            manifest['slides'].append({'slide': index, 'title': data['title'],
                'background_file': asset['file'], 'background_origin': asset['origin'],
                'animated_shape_ids': [s.Id for s in animated],
                'static_text': [{'shape_id': footer.Id, 'reason': 'Persistent page orientation marker'}],
                'transition': 'fade', 'estimated_talk_seconds': data['talk_seconds']})
        pres.SaveAs(str(deck_path.resolve()), 24)
        pres.SaveAs(str((dest / 'sample.pdf').resolve()), 32)
        # SaveAs(PDF) does not change the open document's editable slide content.
        preview = dest / 'preview'
        preview.mkdir(exist_ok=True)
        for i in range(1, pres.Slides.Count + 1):
            pres.Slides.Item(i).Export(str((preview / f'slide-{i:02}.png').resolve()), 'PNG', 1600, 900)
        print(variant, 'PPTX/PDF/previews created', flush=True)
        if make_video:
            # Use identical native effects with automatic triggers for a short preview.
            for i in range(1, pres.Slides.Count + 1):
                slide = pres.Slides.Item(i)
                for j in range(1, slide.TimeLine.MainSequence.Count + 1):
                    timing = slide.TimeLine.MainSequence.Item(j).Timing
                    timing.TriggerType = 2
                    timing.TriggerDelayTime = 0.2 + (j - 1) * 0.65
                slide.SlideShowTransition.AdvanceOnClick = 0
                slide.SlideShowTransition.AdvanceOnTime = -1
                slide.SlideShowTransition.AdvanceTime = 6.0
            video = dest / 'animation-preview.mp4'
            pres.CreateVideo(str(video.resolve()), True, 6, 720, 24, 75)
            start = time.monotonic()
            while int(pres.CreateVideoStatus) in (0, 1, 2):
                if time.monotonic() - start > 300:
                    raise TimeoutError('PowerPoint video export exceeded 300 seconds')
                pythoncom.PumpWaitingMessages()
                time.sleep(1)
            if int(pres.CreateVideoStatus) != 3 or not video.exists():
                raise RuntimeError('PowerPoint video export failed')
            print(variant, 'native animation video created', video.stat().st_size, flush=True)
        pres.Saved = -1  # Preview trigger changes must not overwrite manual PPTX.
        (dest / 'build-record.json').write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
        return manifest
    finally:
        pres.Close()
        app.Quit()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--variant', choices=['chatgpt-reading', 'deepseek-nature', 'all'], default='all')
    parser.add_argument('--output-root', type=Path, required=True)
    parser.add_argument('--video', action='store_true')
    args = parser.parse_args()
    pythoncom.CoInitialize()
    try:
        variants = ['chatgpt-reading', 'deepseek-nature'] if args.variant == 'all' else [args.variant]
        for variant in variants:
            build(variant, args.output_root.resolve(), args.video)
    finally:
        pythoncom.CoUninitialize()


if __name__ == '__main__':
    main()
