#!/usr/bin/env python3
"""
google_docs_high_fidelity_export.py

High-Fidelity Exporter for Google Docs - Python version

Usage:
  python google_docs_high_fidelity_export.py <DOC_ID> <OUTPUT_DIR>

Example:
  python google_docs_high_fidelity_export.py 1AbCdEfgHIjKLMnOP /path/to/export-dir

Dependencies:
  pip install google-auth google-auth-httplib2 google-auth-oauthlib google-api-python-client

This script attempts to replicate the "high-fidelity" HTML export features found in the original
Node.js sample, including heading styles, line spacing, images, lists, TOC, tables, etc.
"""

import os
import sys
import json
import math
import base64
from pathlib import Path

# Google API client libraries
from google.oauth2 import service_account
from googleapiclient.discovery import build

# ------------- CONFIG -------------
SERVICE_ACCOUNT_KEY_FILE = "service_account.json"  # Path to your service account JSON key
EMBED_IMAGES_AS_BASE64 = False  # If True, images are embedded directly as base64 data URIs

# Basic alignment map for LTR paragraphs
ALIGNMENT_MAP_LTR = {
    "START": "left",
    "CENTER": "center",
    "END": "right",
    "JUSTIFIED": "justify"
}


def main():
    if len(sys.argv) < 3:
        print("Usage: python google_docs_high_fidelity_export.py <DOC_ID> <OUTPUT_DIR>")
        sys.exit(1)

    doc_id = sys.argv[1]
    output_dir = sys.argv[2]

    export_doc_to_html(doc_id, output_dir)


def export_doc_to_html(doc_id, output_dir):
    """
    Main entry point to export a Google Doc to a high-fidelity HTML file (index.html).
    Images are saved in an 'images/' subdirectory.
    """
    Path(output_dir).mkdir(parents=True, exist_ok=True)
    images_dir = os.path.join(output_dir, "images")
    Path(images_dir).mkdir(parents=True, exist_ok=True)

    # Auth & fetch doc
    auth_client = get_auth_client()
    docs_service = build("docs", "v1", credentials=auth_client)
    doc = docs_service.documents().get(documentId=doc_id).execute()
    print(f"Exporting doc: {doc.get('title')}")

    # Build named styles map
    named_styles_map = build_named_styles_map(doc)

    # Container width from docStyle
    container_px = compute_doc_container_width(doc)

    # Build global CSS
    global_css = generate_global_css(doc, container_px)

    used_fonts = set()
    html_lines = []

    # Basic HTML skeleton
    html_lines.append("<!DOCTYPE html>")
    html_lines.append('<html lang="en">')
    html_lines.append("<head>")
    html_lines.append('  <meta charset="UTF-8">')
    html_lines.append('  <meta name="viewport" content="width=device-width">')
    html_lines.append(f"  <title>{escape_html(doc.get('title', ''))}</title>")
    html_lines.append("  <style>")
    html_lines.append(global_css)
    html_lines.append("  </style>")
    html_lines.append("</head>")
    html_lines.append("<body>")
    html_lines.append('<div class="doc-content">')

    list_stack = []
    body_content = doc.get("body", {}).get("content", [])

    for element in body_content:
        if "sectionBreak" in element:
            html_lines.append('<div class="section-break"></div>')
            continue

        if "tableOfContents" in element:
            # Close any open lists
            close_all_lists(list_stack, html_lines)
            toc_html = render_table_of_contents(
                element["tableOfContents"],
                doc,
                used_fonts,
                auth_client,
                output_dir,
                named_styles_map
            )
            html_lines.append(toc_html)
            continue

        if "paragraph" in element:
            paragraph_result = render_paragraph(
                element["paragraph"],
                doc,
                used_fonts,
                list_stack,
                auth_client,
                output_dir,
                images_dir,
                named_styles_map
            )
            html = paragraph_result["html"]
            list_change = paragraph_result["listChange"]

            if list_change:
                handle_list_state(list_change, list_stack, html_lines)

            if list_stack:
                # If we are inside a list, wrap paragraph in <li>
                html_lines.append(f"<li>{html}</li>")
            else:
                html_lines.append(html)
            continue

        if "table" in element:
            close_all_lists(list_stack, html_lines)
            table_html = render_table(
                element["table"],
                doc,
                used_fonts,
                auth_client,
                output_dir,
                images_dir,
                named_styles_map
            )
            html_lines.append(table_html)
            continue

    # Close any remaining open lists
    close_all_lists(list_stack, html_lines)

    html_lines.append("</div>")
    html_lines.append("</body>")
    html_lines.append("</html>")

    # Insert Google Fonts if needed
    font_link = build_google_fonts_link(list(used_fonts))
    if font_link:
        # Insert link after </title>
        for i, line in enumerate(html_lines):
            if "</title>" in line:
                html_lines.insert(i + 1, f'  <link rel="stylesheet" href="{font_link}">')
                break

    # Write index.html
    index_path = os.path.join(output_dir, "index.html")
    with open(index_path, "w", encoding="utf-8") as f:
        f.write("\n".join(html_lines))

    print(f"HTML exported to: {index_path}")


# -----------------------------------------------------
# Column width from doc documentStyle
# -----------------------------------------------------
def compute_doc_container_width(doc):
    container_px = 800  # fallback
    ds = doc.get("documentStyle", {})
    page_size = ds.get("pageSize", {})
    page_w = page_size.get("width", {}).get("magnitude", None)
    if page_w:
        left_m = ds.get("marginLeft", {}).get("magnitude", 72)
        right_m = ds.get("marginRight", {}).get("magnitude", 72)
        usable_pts = page_w - (left_m + right_m)
        if usable_pts > 0:
            container_px = pt_to_px(usable_pts)
    # small tweak
    container_px += 64
    return container_px


# -----------------------------------------------------
# Global CSS with heading overrides
# -----------------------------------------------------
def generate_global_css(doc, container_px):
    lines = []
    lines.append("""
/* Reset heading sizes so doc-based inline style rules. */
h1, h2, h3, h4, h5, h6 {
  margin: 1em 0;
  font-size: 1em;
  font-weight: normal;
}

body {
  font-family: sans-serif;
}
.doc-content {
  margin: 1em auto;
  max-width: """ + str(container_px) + """px;
  padding: 2em 1em;
}
p, li {
  margin: 0.5em 0;
}
img {
  display: inline-block;
  max-width: 100%;
  height: auto;
}
.bold { font-weight: bold; }
.italic { font-style: italic; }
.underline { text-decoration: underline; }
.strikethrough { text-decoration: line-through; }
.superscript {
  vertical-align: super;
  font-size: 0.8em;
}
.subscript {
  vertical-align: sub;
  font-size: 0.8em;
}
.section-break {
  page-break-before: always;
}

.doc-toc {
  margin: 0.5em 0;
  padding: 0.5em;
}

.subtitle {
  display: block;
  white-space: pre-wrap;
}
.doc-table {
  border-collapse: collapse;
  margin: 0.5em 0;
}
""")

    doc_style = doc.get("documentStyle", {})
    page_size = doc_style.get("pageSize", {})
    width_pts = page_size.get("width", {}).get("magnitude")
    height_pts = page_size.get("height", {}).get("magnitude")
    if width_pts and height_pts:
        top_m = doc_style.get("marginTop", {}).get("magnitude", 72)
        right_m = doc_style.get("marginRight", {}).get("magnitude", 72)
        bot_m = doc_style.get("marginBottom", {}).get("magnitude", 72)
        left_m = doc_style.get("marginLeft", {}).get("magnitude", 72)
        w_in = width_pts / 72
        h_in = height_pts / 72
        top_in = top_m / 72
        right_in = right_m / 72
        bot_in = bot_m / 72
        left_in = left_m / 72

        lines.append(f"""
@page {{
  size: {w_in}in {h_in}in;
  margin: {top_in}in {right_in}in {bot_in}in {left_in}in;
}}
""")

    return "\n".join(lines)


# -----------------------------------------------------
# Table of Contents (Indentation by heading level)
# -----------------------------------------------------
def render_table_of_contents(toc, doc, used_fonts, auth_client, output_dir, named_styles_map):
    html = '<div class="doc-toc">\n'
    content = toc.get("content", [])
    for c in content:
        paragraph = c.get("paragraph")
        if not paragraph:
            continue

        heading_level = 1
        for elem in paragraph.get("elements", []):
            st = elem.get("textRun", {}).get("textStyle", {})
            link_obj = st.get("link", {})
            if "headingId" in link_obj:
                lvl = find_heading_level_by_id(doc, link_obj["headingId"])
                if lvl > heading_level:
                    heading_level = lvl
        # clamp heading level between 1..4 for TOC styling
        if heading_level < 1:
            heading_level = 1
        if heading_level > 4:
            heading_level = 4

        paragraph_result = render_paragraph(
            paragraph,
            doc,
            used_fonts,
            [],  # no list stack for TOC
            auth_client,
            output_dir,
            None,
            named_styles_map
        )
        p_html = paragraph_result["html"]
        html += f'<div class="toc-level-{heading_level}">{p_html}</div>\n'

    html += "</div>\n"
    return html


def find_heading_level_by_id(doc, heading_id):
    content = doc.get("body", {}).get("content", [])
    for e in content:
        paragraph = e.get("paragraph")
        if paragraph:
            ps = paragraph.get("paragraphStyle", {})
            if ps.get("headingId") == heading_id:
                named = ps.get("namedStyleType", "NORMAL_TEXT")
                if named.startswith("HEADING_"):
                    lv = int(named.replace("HEADING_", ""), 10)
                    if 1 <= lv <= 6:
                        return lv
    return 1


# -----------------------------------------------------
# Paragraph
# -----------------------------------------------------
def render_paragraph(paragraph, doc, used_fonts, list_stack,
                     auth_client, output_dir, images_dir, named_styles_map):
    style = paragraph.get("paragraphStyle", {})
    named_type = style.get("namedStyleType", "NORMAL_TEXT")

    # Merge doc-based style (paragraphStyle + textStyle) from the named style
    merged_para_style = {}
    merged_text_style = {}
    if named_type in named_styles_map:
        merged_para_style = deep_copy(named_styles_map[named_type]["paragraphStyle"])
        merged_text_style = deep_copy(named_styles_map[named_type]["textStyle"])
    deep_merge(merged_para_style, style)

    # bullet logic
    is_rtl = (merged_para_style.get("direction") == "RIGHT_TO_LEFT")
    list_change = None
    bullet = paragraph.get("bullet")
    if bullet:
        list_change = detect_list_change(bullet, doc, list_stack, is_rtl)
    else:
        if list_stack:
            top = list_stack[-1]
            list_change = f"end{top.upper()}"

    # Title => <h1 class="title">, Subtitle => <h2 class="subtitle">, heading => <hX>, else <p>
    tag = "p"
    if named_type == "TITLE":
        tag = 'h1 class="title"'
    elif named_type == "SUBTITLE":
        tag = 'h2 class="subtitle"'
    elif named_type.startswith("HEADING_"):
        lv = int(named_type.replace("HEADING_", ""))
        if 1 <= lv <= 6:
            tag = f"h{lv}"

    heading_id_attr = ""
    if "headingId" in style:
        heading_id_attr = f' id="heading-{escape_html(style["headingId"])}"'

    # alignment flipping
    align = merged_para_style.get("alignment")
    if is_rtl and align in ("START", "END"):
        align = "END" if align == "START" else "START"

    inline_style = ""
    if align and align in ALIGNMENT_MAP_LTR:
        inline_style += f"text-align:{ALIGNMENT_MAP_LTR[align]};"

    # line spacing
    line_spacing = merged_para_style.get("lineSpacing")
    if line_spacing:
        # browsers default ~1.2; multiplied by 1.25 for a closer match
        ls = (line_spacing * 1.25) / 100.0
        inline_style += f"line-height:{ls};"

    # spaceAbove / spaceBelow
    space_above = merged_para_style.get("spaceAbove", {}).get("magnitude")
    if space_above:
        inline_style += f"margin-top:{pt_to_px(space_above)}px;"
    space_below = merged_para_style.get("spaceBelow", {}).get("magnitude")
    if space_below:
        inline_style += f"margin-bottom:{pt_to_px(space_below)}px;"

    # indent
    indent_first_line = merged_para_style.get("indentFirstLine", {}).get("magnitude")
    if indent_first_line:
        inline_style += f"text-indent:{pt_to_px(indent_first_line)}px;"
    else:
        indent_start = merged_para_style.get("indentStart", {}).get("magnitude")
        if indent_start:
            inline_style += f"margin-left:{pt_to_px(indent_start)}px;"

    dir_attr = ""
    if is_rtl:
        dir_attr = ' dir="rtl"'

    # Merge text runs
    merged_runs = merge_text_runs(paragraph.get("elements", []))
    inner_html = ""

    for r in merged_runs:
        if "inlineObjectElement" in r:
            obj_id = r["inlineObjectElement"]["inlineObjectId"]
            inner_html += render_inline_object(obj_id, doc, auth_client, output_dir, images_dir)
        elif "textRun" in r:
            inner_html += render_text_run(r["textRun"], used_fonts, merged_text_style)

    # Construct final paragraph HTML
    # tag might be 'h1 class="title"', so extract the actual tag name for closing
    tag_parts = tag.split(" ", 1)
    tag_name = tag_parts[0] if tag_parts else "p"
    paragraph_html = f"<{tag}{heading_id_attr}{dir_attr}>{inner_html}</{tag_name}>"

    if inline_style:
        close_bracket = paragraph_html.find(">")
        if close_bracket > 0:
            paragraph_html = (
                paragraph_html[:close_bracket]
                + f' style="{inline_style}"'
                + paragraph_html[close_bracket:]
            )

    return {"html": paragraph_html, "listChange": list_change}


# -----------------------------------------------------
# Merging text runs
# -----------------------------------------------------
def merge_text_runs(elements):
    merged = []
    last = None

    for e in elements:
        if "inlineObjectElement" in e:
            # This is an image or similar object
            merged.append({"inlineObjectElement": e["inlineObjectElement"]})
            last = None
        elif "textRun" in e:
            style = e["textRun"].get("textStyle", {})
            content = e["textRun"].get("content", "")
            if last and "textRun" in last:
                # compare style
                if is_same_text_style(last["textRun"]["textStyle"], style):
                    # merge content
                    last["textRun"]["content"] += content
                else:
                    merged.append({"textRun": {"content": content, "textStyle": deep_copy(style)}})
                    last = merged[-1]
            else:
                merged.append({"textRun": {"content": content, "textStyle": deep_copy(style)}})
                last = merged[-1]

    return merged


def is_same_text_style(a, b):
    fields = [
        "bold", "italic", "underline", "strikethrough",
        "baselineOffset", "fontSize", "weightedFontFamily",
        "foregroundColor", "link"
    ]
    for f in fields:
        if json.dumps(a.get(f)) != json.dumps(b.get(f)):
            return False
    return True


# -----------------------------------------------------
# Rendering text runs
# -----------------------------------------------------
def render_text_run(text_run, used_fonts, base_style):
    final_style = deep_copy(base_style or {})
    deep_merge(final_style, text_run.get("textStyle", {}))

    content = text_run.get("content", "")
    # doc strings often have trailing \n
    content = content.rstrip("\n")

    css_classes = []
    inline_style = ""

    if final_style.get("bold"):
        css_classes.append("bold")
    if final_style.get("italic"):
        css_classes.append("italic")
    if final_style.get("underline"):
        css_classes.append("underline")
    if final_style.get("strikethrough"):
        css_classes.append("strikethrough")

    baseline_offset = final_style.get("baselineOffset")
    if baseline_offset == "SUPERSCRIPT":
        css_classes.append("superscript")
    elif baseline_offset == "SUBSCRIPT":
        css_classes.append("subscript")

    font_size = final_style.get("fontSize", {}).get("magnitude")
    if font_size:
        inline_style += f"font-size:{font_size}pt;"

    weighted_font_family = final_style.get("weightedFontFamily", {})
    if "fontFamily" in weighted_font_family:
        fam = weighted_font_family["fontFamily"]
        used_fonts.add(fam)
        inline_style += f"font-family:'{fam}', sans-serif;"

    fg_color = final_style.get("foregroundColor", {}).get("color", {}).get("rgbColor")
    if fg_color:
        hex_val = rgb_to_hex(
            fg_color.get("red", 0),
            fg_color.get("green", 0),
            fg_color.get("blue", 0)
        )
        inline_style += f"color:{hex_val};"

    open_tag = "<span"
    if css_classes:
        open_tag += f' class="{" ".join(css_classes)}"'
    if inline_style:
        open_tag += f' style="{inline_style}"'
    open_tag += ">"
    close_tag = "</span>"

    link_obj = final_style.get("link")
    if link_obj:
        link_href = ""
        if "headingId" in link_obj:
            link_href = f"#heading-{escape_html(link_obj['headingId'])}"
        elif "url" in link_obj:
            link_href = link_obj["url"]
        if link_href:
            open_tag = f'<a href="{escape_html(link_href)}"'
            if css_classes:
                open_tag += f' class="{" ".join(css_classes)}"'
            if inline_style:
                open_tag += f' style="{inline_style}"'
            open_tag += ">"
            close_tag = "</a>"

    return open_tag + escape_html(content) + close_tag


# -----------------------------------------------------
# Inline Objects (Images)
# -----------------------------------------------------
def render_inline_object(object_id, doc, auth_client, output_dir, images_dir):
    inline_obj = doc.get("inlineObjects", {}).get(object_id)
    if not inline_obj:
        return ""

    embedded = inline_obj.get("inlineObjectProperties", {}).get("embeddedObject")
    if not embedded or "imageProperties" not in embedded:
        return ""

    image_props = embedded["imageProperties"]
    content_uri = image_props.get("contentUri")
    size = image_props.get("size", {})

    scale_x = embedded.get("transform", {}).get("scaleX", 1)
    scale_y = embedded.get("transform", {}).get("scaleY", 1)

    if not content_uri:
        return ""

    # Fetch the image content
    base64_data = fetch_as_base64(content_uri, auth_client)
    img_data = base64.b64decode(base64_data)

    file_name = f"image_{object_id}.png"
    file_path = os.path.join(images_dir, file_name)

    # Write the file
    with open(file_path, "wb") as f:
        f.write(img_data)

    # Build the <img> HTML
    alt = embedded.get("title", "") or embedded.get("description", "")
    style = ""
    width_pts = size.get("width", {}).get("magnitude")
    height_pts = size.get("height", {}).get("magnitude")

    if width_pts and height_pts:
        w_px = round(width_pts * 1.3333 * scale_x)
        h_px = round(height_pts * 1.3333 * scale_y)
        style = f"max-width:{w_px}px; max-height:{h_px}px;"

    img_src = os.path.relpath(file_path, output_dir)

    return f'<img src="{escape_html(img_src)}" alt="{escape_html(alt)}" style="{style}" />'


# -----------------------------------------------------
# Lists
# -----------------------------------------------------
def detect_list_change(bullet, doc, list_stack, is_rtl):
    list_id = bullet["listId"]
    nesting_level = bullet.get("nestingLevel", 0)
    list_def = doc.get("lists", {}).get(list_id)
    if not list_def:
        return None

    nesting_levels = list_def.get("listProperties", {}).get("nestingLevels", [])
    if nesting_level >= len(nesting_levels):
        return None

    glyph = nesting_levels[nesting_level]
    glyph_type = glyph.get("glyphType", "").lower()
    is_numbered = "number" in glyph_type

    current_type = "ol" if is_numbered else "ul"
    rtl_flag = "_RTL" if is_rtl else ""

    if list_stack:
        top = list_stack[-1]
        if not top.startswith(current_type):
            # switch from ul->ol or vice versa
            return f"end{top.upper()}|start{current_type.upper()}{rtl_flag}"
        return None
    else:
        return f"start{current_type.upper()}{rtl_flag}"


def handle_list_state(list_change, list_stack, html_lines):
    actions = list_change.split("|")
    for action in actions:
        if action.startswith("start"):
            if "UL_RTL" in action:
                html_lines.append('<ul dir="rtl">')
                list_stack.append("ul_rtl")
            elif "OL_RTL" in action:
                html_lines.append('<ol dir="rtl">')
                list_stack.append("ol_rtl")
            elif "UL" in action:
                html_lines.append("<ul>")
                list_stack.append("ul")
            else:
                html_lines.append("<ol>")
                list_stack.append("ol")
        elif action.startswith("end"):
            top = list_stack.pop() if list_stack else None
            if top and top.startswith("u"):
                html_lines.append("</ul>")
            else:
                html_lines.append("</ol>")


def close_all_lists(list_stack, html_lines):
    while list_stack:
        top = list_stack.pop()
        if top.startswith("u"):
            html_lines.append("</ul>")
        else:
            html_lines.append("</ol>")


# -----------------------------------------------------
# Table
# -----------------------------------------------------
def render_table(table, doc, used_fonts, auth_client, output_dir, images_dir, named_styles_map):
    html = '<table class="doc-table" style="border-collapse:collapse; border:1px solid #ccc;">'
    for row in table.get("tableRows", []):
        html += "<tr>"
        for cell in row.get("tableCells", []):
            html += '<td style="border:1px solid #ccc; padding:0.5em;">'
            for content in cell.get("content", []):
                paragraph = content.get("paragraph")
                if paragraph:
                    paragraph_result = render_paragraph(
                        paragraph,
                        doc,
                        used_fonts,
                        [],
                        auth_client,
                        output_dir,
                        images_dir,
                        named_styles_map
                    )
                    html += paragraph_result["html"]
            html += "</td>"
        html += "</tr>"
    html += "</table>"
    return html


# -----------------------------------------------------
# Named Styles
# -----------------------------------------------------
def build_named_styles_map(doc):
    named = doc.get("namedStyles", {}).get("styles", [])
    style_map = {}
    for s in named:
        style_map[s["namedStyleType"]] = {
            "paragraphStyle": s.get("paragraphStyle", {}),
            "textStyle": s.get("textStyle", {})
        }
    return style_map


# -----------------------------------------------------
# AUTH, UTILITIES
# -----------------------------------------------------
def get_auth_client():
    """
    Create an authenticated client from a service account JSON key.
    """
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_KEY_FILE,
        scopes=[
            "https://www.googleapis.com/auth/documents.readonly",
            "https://www.googleapis.com/auth/drive.readonly",
        ]
    )
    return creds


def fetch_as_base64(url, auth_client):
    """
    Fetch image data with authentication, return base64 encoded string.
    """
    import requests
    token = auth_client.token
    if not token:
        auth_client.refresh(requests.Request())
        token = auth_client.token

    # Make an authorized request
    headers = {"Authorization": f"Bearer {token}"}
    resp = requests.get(url, headers=headers)
    resp.raise_for_status()
    return base64.b64encode(resp.content).decode("utf-8")


def deep_merge(base, overlay):
    """
    Deep merge dictionary 'overlay' into 'base'.
    """
    for k, v in overlay.items():
        if isinstance(v, dict) and v and not isinstance(v, list):
            base.setdefault(k, {})
            deep_merge(base[k], v)
        else:
            base[k] = v


def deep_copy(obj):
    return json.loads(json.dumps(obj))


def escape_html(text):
    return (
        text.replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;")
    )


def pt_to_px(pts):
    return round(pts * 1.3333)


def rgb_to_hex(r, g, b):
    nr = round(r * 255)
    ng = round(g * 255)
    nb = round(b * 255)
    return "#" + "".join([f"{x:02x}" for x in (nr, ng, nb)])


def build_google_fonts_link(font_families):
    if not font_families:
        return ""
    unique = list(set(font_families))
    # E.g.: "Roboto" => "Roboto"; multiple => &family=Open+Sans&family=Roboto
    families_param = "&family=".join(f.replace(" ", "+") for f in unique)
    return f"https://fonts.googleapis.com/css2?family={families_param}&display=swap"


if __name__ == "__main__":
    main()

