from flask import Flask, render_template_string, request, send_file
from werkzeug.utils import secure_filename
from pathlib import Path
import tempfile
import os
import io

# Reuse logic from existing module
from style_tableau_pptx import (
    Presentation,
    Inches,
    Pt,
    apply_title_case,
    guess_title_from_slide,
    extract_title_from_image,
    find_layout,
    add_slide_number,
    add_rounded_corners,
    fit_image_on_blank,
    apply_box_shadow,
)
import easyocr

app = Flask(__name__)

# Serve hero image via /hero.png (expects hero_image.png alongside this script)
@app.get('/hero.png')
def hero_png():
    hero_path = Path(__file__).with_name('hero_image.png')
    if not hero_path.exists():
        return ('Hero image not found', 404)
    return send_file(str(hero_path), mimetype='image/png')

FORM = """
<!doctype html>
<html>
  <head>
    <meta charset=\"utf-8\" />
    <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\" />
    <title>Style & Combine PPTX</title>
    <style>
      :root {
        --bg: #f7f9fc;
        --panel: #ffffff;
        --muted: #667085;
        --text: #0f172a;
        --accent: #2563eb; /* blue-600 */
        --accent-2: #60a5fa; /* blue-400 */
        --border: #e5e7eb;
      }
      * { box-sizing: border-box; }
      html, body { height: 100%; scroll-behavior: smooth; }
      body {
        margin: 0; font: 14px/1.5 system-ui, -apple-system, Segoe UI, Roboto, sans-serif; color: var(--text);
        background: radial-gradient(1200px 600px at 20% -10%, rgba(96,165,250,0.25), transparent 60%),
                    radial-gradient(1000px 600px at 80% 0%, rgba(37,99,235,0.18), transparent 60%),
                    var(--bg);
      }
      .header {
        padding: 28px 20px; background: linear-gradient(135deg, rgba(37,99,235,0.12), rgba(96,165,250,0.08));
        border-bottom: 1px solid var(--border);
      }
      .header h1 { margin: 0; font-size: 22px; letter-spacing: 0.3px; }
      .hero { max-width: 1100px; margin: 0 auto; padding: 24px 16px 18px; display: grid; grid-template-columns: 1.1fr 1fr; gap: 24px; align-items: center; min-height: 320px; }
      .hero-left h1 { font-size: 40px; line-height: 1.1; margin: 0 0 10px; }
      .subtitle { color: var(--muted); margin: 0 0 16px; font-size: 15px; }
      .cta { display: none; }
      .hero-art { height: 280px; border-radius: 18px; border: 1px solid var(--border); background: #f0f5ff; overflow: hidden;
        box-shadow: 0 8px 24px rgba(37,99,235,0.12), 0 2px 8px rgba(15,23,42,0.06);
      }
      .hero-art img { width: 100%; height: 100%; object-fit: cover; display: block; }
      .container { max-width: 1100px; margin: 22px auto; padding: 0 16px; }
      .grid { display: grid; grid-template-columns: 1.2fr 1fr; gap: 16px; }
      .card { background: var(--panel); border: 1px solid var(--border); border-radius: 12px; padding: 14px; box-shadow: 0 1px 2px rgba(2,6,23,0.04); }
      .card h3 { margin: 0 0 8px; font-size: 14px; color: #1e293b; letter-spacing: 0.2px; }
      .section { margin-top: 10px; }
      .stack { display: grid; gap: 10px; }
      label { display: block; margin-bottom: 4px; color: var(--muted); font-size: 12px; }
      input[type=file] { display: none; }
      .drop { border: 1px dashed var(--border); border-radius: 10px; padding: 18px; text-align: center; color: var(--muted); background: #f3f6fb; }
      .drop.drag { border-color: var(--accent); color: var(--text); background: #eaf2ff; }
      .btn { appearance: none; border: 1px solid var(--border); border-radius: 10px; padding: 10px 14px; background: #ffffff; color: var(--text); cursor: pointer; }
      .btn:hover { border-color: #cbd5e1; background: #f8fafc; }
      .btn.primary { background: linear-gradient(135deg, #3b82f6, #2563eb); border: none; color: #ffffff; font-weight: 700; }
      .btn.primary:hover { filter: brightness(0.98); }
      .row { display: grid; grid-template-columns: repeat(2, 1fr); gap: 10px; }
      .row-3 { display: grid; grid-template-columns: repeat(3, 1fr); gap: 10px; }
      .row-4 { display: grid; grid-template-columns: repeat(4, 1fr); gap: 10px; }
      .input, select { width: 100%; padding: 10px 12px; border-radius: 10px; border: 1px solid var(--border); background: #ffffff; color: var(--text); }
      .muted { color: var(--muted); font-size: 12px; }
      .file-list { list-style: none; margin: 0; padding: 0; display: grid; gap: 8px; }
      .file-item { display: grid; grid-template-columns: 1fr auto; gap: 8px; align-items: center; padding: 8px 10px; border: 1px solid var(--border); border-radius: 8px; background: #ffffff; }
      .file-item[draggable=true] { cursor: grab; }
      .pill { padding: 2px 8px; border-radius: 999px; background: #eff6ff; border: 1px solid #bfdbfe; font-size: 12px; color: #1e3a8a; }
      .footer { margin-top: 16px; display: flex; gap: 10px; align-items: center; }
      .status { min-height: 24px; }
      .kbd { font-family: ui-monospace, SFMono-Regular, Menlo, monospace; font-size: 12px; background: #f8fafc; border: 1px solid var(--border); padding: 2px 6px; border-radius: 6px; }
      .color { height: 36px; width: 100%; border-radius: 10px; border: 1px solid var(--border); background: transparent; }
      @media (max-width: 980px) {
        .grid { grid-template-columns: 1fr; }
        .hero { grid-template-columns: 1fr; padding-bottom: 8px; }
        .hero-art { height: 220px; }
        .hero-left h1 { font-size: 32px; }
      }
    </style>
  </head>
  <body>
    <div class=\"header\">
      <div class=\"hero\">
        <div class=\"hero-left\">
          <h1>Quickly style your Tableau PowerPoint exports</h1>
          <p class=\"subtitle\">Drop in Tableau-exported decks, apply brand settings, and download a polished, on‑brand presentation.</p>
        </div>
        <div class=\"hero-right\">
          <div class=\"hero-art\"><img src=\"/hero.png\" alt=\"Hero graphic\"/></div>
        </div>
      </div>
    </div>
    <div class=\"container\" id=\"app\">
      <div class=\"grid\">
        <div class=\"card\">
          <h3>Files</h3>
          <div class=\"section stack\">
            <div>
              <label>Template PPTX</label>
              <div id=\"tplDrop\" class=\"drop\">
                <div><strong>Drop template</strong> or <button class=\"btn\" type=\"button\" id=\"tplPick\">Browse</button></div>
                <div class=\"muted\">Accepted: .pptx</div>
              </div>
              <input id=\"tplInput\" type=\"file\" accept=\".pptx\" />
              <div id=\"tplName\" class=\"muted\"></div>
            </div>
            <div>
              <label>Input PPTX (order matters)</label>
              <div id=\"inDrop\" class=\"drop\">
                <div><strong>Drop decks</strong> or <button class=\"btn\" type=\"button\" id=\"inPick\">Browse</button></div>
                <div class=\"muted\">Drag to reorder. Accepted: .pptx</div>
              </div>
              <input id=\"inInput\" type=\"file\" accept=\".pptx\" multiple />
              <ul id=\"fileList\" class=\"file-list\"></ul>
            </div>
          </div>
        </div>

        <div class=\"card\">
          <h3>Options</h3>
          <div class=\"stack\">
            <div class=\"row\">
              <div>
                <label>Title Case</label>
                <select id=\"title_case\" class=\"input\">
                  <option value=\"smart\" selected>smart</option>
                  <option value=\"camel\">camel</option>
                  <option value=\"upper\">upper</option>
                  <option value=\"lower\">lower</option>
                </select>
              </div>
              <div>
                <label>Title Font Size</label>
                <input id=\"title_font_size\" class=\"input\" type=\"number\" value=\"28\" min=\"10\" max=\"80\" />
              </div>
            </div>

            <div class=\"row\">
              <div>
                <label>Border Radius (px)</label>
                <input id=\"border_radius\" class=\"input\" type=\"number\" value=\"10\" min=\"0\" max=\"40\" />
              </div>
              <div>
                <label>Shadow Enabled</label>
                <select id=\"shadow\" class=\"input\"><option value=\"on\" selected>on</option><option value=\"off\">off</option></select>
              </div>
            </div>

            <div class=\"row-4\">
              <div>
                <label>Shadow Color</label>
                <input id=\"shadow_color\" class=\"color\" type=\"color\" value=\"#000000\" />
              </div>
              <div>
                <label>Transparency</label>
                <input id=\"shadow_transparency\" class=\"input\" type=\"number\" step=\"0.05\" value=\"0.8\" />
              </div>
              <div>
                <label>Blur (pt)</label>
                <input id=\"shadow_blur\" class=\"input\" type=\"number\" value=\"15\" />
              </div>
              <div>
                <label>Angle (deg)</label>
                <input id=\"shadow_angle\" class=\"input\" type=\"number\" value=\"34\" />
              </div>
            </div>
            <div class=\"row-3\">
              <div>
                <label>Distance (pt)</label>
                <input id=\"shadow_distance\" class=\"input\" type=\"number\" value=\"3\" />
              </div>
              <div>
                <label>Image Left (in)</label>
                <input id=\"image_left\" class=\"input\" type=\"number\" step=\"0.1\" value=\"2.5\" />
              </div>
              <div>
                <label>Image Top (in)</label>
                <input id=\"image_top\" class=\"input\" type=\"number\" step=\"0.1\" value=\"1.7\" />
              </div>
            </div>

            <div class=\"footer\">
              <button id=\"processBtn\" class=\"btn primary\" type=\"button\">Process</button>
              <span id=\"status\" class=\"status muted\"></span>
            </div>
          </div>
        </div>
      </div>
      <div id=\"downloadArea\" class=\"section\"></div>
    </div>

    <script>
      const tplInput = document.getElementById('tplInput');
      const tplDrop = document.getElementById('tplDrop');
      const tplPick = document.getElementById('tplPick');
      const tplName = document.getElementById('tplName');
      const inInput = document.getElementById('inInput');
      const inDrop = document.getElementById('inDrop');
      const inPick = document.getElementById('inPick');
      const listEl = document.getElementById('fileList');
      const processBtn = document.getElementById('processBtn');
      const statusEl = document.getElementById('status');
      const downloadArea = document.getElementById('downloadArea');

      let templateFile = null;
      let inputFiles = [];

      function fmtSize(bytes){
        if (!bytes && bytes !== 0) return '';
        const units=['B','KB','MB','GB'];
        let i=0; let v=bytes; while(v>1024 && i<units.length-1){v/=1024;i++;}
        return v.toFixed(1)+' '+units[i];
      }

      function renderList(){
        listEl.innerHTML = '';
        inputFiles.forEach((f, idx) => {
          const li = document.createElement('li');
          li.className = 'file-item';
          li.draggable = true;
          li.dataset.index = idx;
          li.innerHTML = `<div><strong>${f.name}</strong> <span class="muted">${fmtSize(f.size)}</span></div><div><button class="btn" data-remove="${idx}">Remove</button></div>`;
          listEl.appendChild(li);
        });
      }

      function handleReorder(ev){
        const src = ev.dataTransfer.getData('text/x-idx');
        const dst = ev.currentTarget.dataset.index;
        if(src==='') return;
        const s = parseInt(src,10); const d = parseInt(dst,10);
        if(Number.isNaN(s) || Number.isNaN(d) || s===d) return;
        const item = inputFiles.splice(s,1)[0];
        inputFiles.splice(d,0,item);
        renderList();
        bindDnD();
      }

      function bindDnD(){
        [...document.querySelectorAll('.file-item')].forEach(el => {
          el.addEventListener('dragstart', e => { e.dataTransfer.setData('text/x-idx', el.dataset.index); });
          el.addEventListener('dragover', e => e.preventDefault());
          el.addEventListener('drop', handleReorder);
        });
        listEl.querySelectorAll('button[data-remove]').forEach(btn => btn.addEventListener('click', e => {
          const idx = parseInt(btn.getAttribute('data-remove'),10);
          inputFiles.splice(idx,1); renderList(); bindDnD();
        }));
      }

      function handleDrop(zone, files){
        zone.classList.remove('drag');
        if(zone===tplDrop){ templateFile = files[0]; tplName.textContent = templateFile ? `${templateFile.name} (${fmtSize(templateFile.size)})` : ''; }
        else { inputFiles.push(...[...files]); renderList(); bindDnD(); }
      }

      function wireDrop(zone, inputEl){
        zone.addEventListener('dragover', e=>{ e.preventDefault(); zone.classList.add('drag'); });
        zone.addEventListener('dragleave', e=> zone.classList.remove('drag'));
        zone.addEventListener('drop', e=>{ e.preventDefault(); handleDrop(zone, e.dataTransfer.files); });
        zone.querySelector('button').addEventListener('click', ()=> inputEl.click());
        inputEl.addEventListener('change', ()=> handleDrop(zone, inputEl.files));
      }

      wireDrop(tplDrop, tplInput);
      wireDrop(inDrop, inInput);

      processBtn.addEventListener('click', async ()=>{
        statusEl.textContent = 'Processing...';
        processBtn.disabled = true;
        downloadArea.innerHTML = '';
        try{
          if(!templateFile) throw new Error('Please select a template PPTX');
          if(inputFiles.length===0) throw new Error('Please add at least one input PPTX');

          const fd = new FormData();
          fd.append('template', templateFile, templateFile.name);
          inputFiles.forEach(f => fd.append('inputs', f, f.name));

          const get = id => document.getElementById(id);
          const colorHex = get('shadow_color').value.replace('#','');
          fd.append('title_case', get('title_case').value);
          fd.append('title_font_size', get('title_font_size').value);
          fd.append('border_radius', get('border_radius').value);
          fd.append('shadow', get('shadow').value === 'on' ? 'on' : 'off');
          fd.append('shadow_color', colorHex);
          fd.append('shadow_transparency', get('shadow_transparency').value);
          fd.append('shadow_blur', get('shadow_blur').value);
          fd.append('shadow_angle', get('shadow_angle').value);
          fd.append('shadow_distance', get('shadow_distance').value);
          fd.append('image_left', get('image_left').value);
          fd.append('image_top', get('image_top').value);

          const resp = await fetch('/', { method: 'POST', body: fd });
          if(!resp.ok) throw new Error('Server error');
          const blob = await resp.blob();
          const url = URL.createObjectURL(blob);
          const a = document.createElement('a');
          a.href = url; a.download = 'styled_output.pptx';
          a.click();
          URL.revokeObjectURL(url);
          statusEl.textContent = 'Done';
          downloadArea.innerHTML = '<span class="pill">Saved: styled_output.pptx</span>';
        } catch(err){
          statusEl.textContent = err.message || 'Error';
        } finally {
          processBtn.disabled = false;
        }
      });
    </script>
  </body>
 </html>
"""

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'GET':
        return render_template_string(FORM)

    # Save uploads to temp dir
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)
        tpl_file = request.files.get('template')
        if not tpl_file:
            return 'Missing template', 400
        tpl_path = tmp / secure_filename(tpl_file.filename)
        tpl_file.save(tpl_path)

        inputs = request.files.getlist('inputs')
        if not inputs:
            return 'No inputs', 400
        input_paths = []
        for f in inputs:
            p = tmp / secure_filename(f.filename)
            f.save(p)
            input_paths.append(p)

        # Options
        title_case = request.form.get('title_case', 'smart')
        title_font_size = int(request.form.get('title_font_size', 28))
        border_radius = int(request.form.get('border_radius', 10))
        shadow_enabled = request.form.get('shadow') == 'on'
        shadow_color_hex = request.form.get('shadow_color', '000000')
        shadow_transparency = float(request.form.get('shadow_transparency', 0.8))
        shadow_blur = int(request.form.get('shadow_blur', 15))
        shadow_angle = int(request.form.get('shadow_angle', 34))
        shadow_distance = int(request.form.get('shadow_distance', 3))
        image_left = float(request.form.get('image_left', 2.5))
        image_top = float(request.form.get('image_top', 1.7))

        try:
            shadow_color = tuple(int(shadow_color_hex[i:i+2], 16) for i in (0,2,4))
        except Exception:
            shadow_color = (0, 0, 0)

        # Build output
        reader = easyocr.Reader(['en'], gpu=False)
        out = Presentation(str(tpl_path))
        tpl = Presentation(str(tpl_path))
        layout = find_layout(tpl)
        add_slide_number(out)

        slide_counter = 0
        for in_path in input_paths:
            src = Presentation(str(in_path))
            for s in src.slides:
                slide_counter += 1
                slide = out.slides.add_slide(layout)

                # Extract first image
                image_path = None
                for shp in s.shapes:
                    if getattr(shp, 'image', None) is not None:
                        ext = shp.image.ext or 'png'
                        tmp_img = tmp / f"_tmp_{slide_counter}.{ext}"
                        with open(tmp_img, 'wb') as fh:
                            fh.write(shp.image.blob)
                        image_path = tmp_img
                        break

                # Title
                title = None
                if image_path and image_path.exists():
                    title = extract_title_from_image(str(image_path), reader)
                if not title:
                    title = guess_title_from_slide(s) or f"Dashboard {slide_counter}"
                title = (title or '')[:120]
                title = apply_title_case(title, title_case)

                if slide.shapes.title:
                    slide.shapes.title.text = title
                    slide.shapes.title.text_frame.paragraphs[0].runs[0].font.size = Pt(title_font_size)
                else:
                    tb = slide.shapes.add_textbox(Inches(0.7), Inches(0.5), out.slide_width - Inches(1.4), Inches(0.6))
                    p = tb.text_frame.paragraphs[0]
                    p.text = title
                    p.runs[0].font.size = Pt(title_font_size)

                # Picture
                if image_path and image_path.exists():
                    add_rounded_corners(str(image_path), radius_px=border_radius)
                    pic = slide.shapes.add_picture(str(image_path), 0, 0)
                    fit_image_on_blank(slide, out, pic, left_in=image_left, top_in=image_top)
                    if shadow_enabled:
                        apply_box_shadow(
                            pic,
                            transparency=shadow_transparency,
                            blur_pt=shadow_blur,
                            angle_deg=shadow_angle,
                            distance_pt=shadow_distance,
                            color=shadow_color,
                        )


        # Save to memory and return
        buf = io.BytesIO()
        out.save(buf)
        buf.seek(0)
        return send_file(buf, as_attachment=True, download_name='styled_output.pptx', mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation')


if __name__ == '__main__':
    # Run server: FLASK_APP=style_pptx_web.py flask run (or python style_pptx_web.py)
    app.run(host='127.0.0.1', port=5001, debug=False)
