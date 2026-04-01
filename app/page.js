'use client';

import { useState } from 'react';
import { parsePptx } from '../lib/parsePptx.js';

export default function Home() {
  const [brandFile, setBrandFile] = useState(null);
  const [form, setForm] = useState({ title: '', description: '', content: '' });
  const [status, setStatus] = useState(null);
  const [errorMsg, setErrorMsg] = useState(null);
  const [previewHtml, setPreviewHtml] = useState(null);
  const [pptxBase64, setPptxBase64] = useState(null);

  function handleChange(e) {
    setForm({ ...form, [e.target.name]: e.target.value });
  }

  async function handleGenerate() {
    if (!brandFile) { setStatus('no-file'); return; }
    if (!form.title.trim() || !form.description.trim() || !form.content.trim()) {
      setStatus('validation'); return;
    }
    setStatus('loading');
    setPreviewHtml(null);
    setPptxBase64(null);
    setErrorMsg(null);

    let colors = [], fonts = [], backgroundImage = null, templateShapes = [];
    try {
      const arrayBuffer = await brandFile.arrayBuffer();
      ({ colors, fonts, backgroundImage, templateShapes } = await parsePptx(arrayBuffer));
    } catch { /* continue */ }

    try {
      const res = await fetch('/api/generate', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ title: form.title, description: form.description, content: form.content, colors, fonts, backgroundImage, templateShapes }),
      });
      const data = await res.json();
      if (!res.ok) { setErrorMsg(data?.error || 'Unknown error'); setStatus('error'); return; }
      setPreviewHtml(data.previewHtml);
      setPptxBase64(data.pptxBase64);
      setStatus('success');
    } catch (e) {
      setErrorMsg(e?.message || 'Network error');
      setStatus('error');
    }
  }

  function handleDownload() {
    if (!pptxBase64) return;
    const bytes = Uint8Array.from(atob(pptxBase64), (c) => c.charCodeAt(0));
    const blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${form.title || 'slide'}.pptx`;
    a.click();
    URL.revokeObjectURL(url);
  }

  return (
    <main style={s.desktop}>
      {/* Window */}
      <div style={s.window}>
        {/* Title Bar */}
        <div style={s.titleBar}>
          <div style={s.titleBarLeft}>
            {/* Tiny icon */}
            <span style={s.titleIcon}>📊</span>
            <span style={s.titleText}>SNV Slide Generator</span>
          </div>
          <div style={s.titleBarButtons}>
            <button style={s.winBtn}>_</button>
            <button style={s.winBtn}>□</button>
            <button style={{ ...s.winBtn, ...s.winBtnClose }}>✕</button>
          </div>
        </div>

        {/* Menu Bar */}
        <div style={s.menuBar}>
          {['File', 'Edit', 'View', 'Help'].map((m) => (
            <span key={m} style={s.menuItem}>{m}</span>
          ))}
        </div>

        {/* Window Body */}
        <div style={s.windowBody}>
          {/* Left panel — form */}
          <div style={s.panel}>
            {/* Group box: Brand File */}
            <fieldset style={s.groupBox}>
              <legend style={s.legend}>Reference File</legend>
              <label style={s.label}>PPTX Template (.pptx):</label>
              <input
                type="file"
                accept=".pptx"
                onChange={(e) => setBrandFile(e.target.files?.[0] || null)}
                style={s.fileInput}
              />
              {brandFile && <p style={s.fileNote}>✔ {brandFile.name}</p>}
            </fieldset>

            {/* Group box: Slide Details */}
            <fieldset style={s.groupBox}>
              <legend style={s.legend}>Slide Details</legend>

              <label style={s.label}>Slide Title:</label>
              <input
                name="title"
                value={form.title}
                onChange={handleChange}
                placeholder="Enter slide title"
                style={s.input}
              />

              <label style={{ ...s.label, marginTop: 8 }}>Short Description:</label>
              <input
                name="description"
                value={form.description}
                onChange={handleChange}
                placeholder="Enter slide description"
                style={s.input}
              />

              <label style={{ ...s.label, marginTop: 8 }}>Slide Content:</label>
              <textarea
                name="content"
                value={form.content}
                onChange={handleChange}
                placeholder="Enter slide content"
                rows={5}
                style={{ ...s.input, resize: 'vertical', height: 90 }}
              />
            </fieldset>

            {/* Status / error messages */}
            {status === 'no-file' && <div style={s.errorBox}>⚠ Please upload a brand PPTX file first.</div>}
            {status === 'validation' && <div style={s.errorBox}>⚠ Please fill in all fields before generating.</div>}
            {status === 'error' && <div style={s.errorBox}>⚠ {errorMsg || 'Something went wrong. Please try again.'}</div>}

            {/* Button row */}
            <div style={s.buttonRow}>
              <button
                onClick={handleGenerate}
                disabled={status === 'loading'}
                style={{ ...s.button, opacity: status === 'loading' ? 0.7 : 1, cursor: status === 'loading' ? 'not-allowed' : 'default' }}
              >
                {status === 'loading' ? 'Generating...' : 'Generate'}
              </button>
              <button style={s.button} onClick={() => { setForm({ title: '', description: '', content: '' }); setBrandFile(null); setStatus(null); setPreviewHtml(null); setPptxBase64(null); }}>
                Reset
              </button>
            </div>
          </div>

          {/* Right panel — preview */}
          <div style={s.previewPanel}>
            <div style={s.previewInset}>
              {previewHtml ? (
                <div style={{ overflowX: 'auto', overflowY: 'auto', width: '100%', height: '100%' }}>
                  <div
                    style={{ transform: 'scale(0.72)', transformOrigin: 'top left', width: 800, height: 450 }}
                    dangerouslySetInnerHTML={{ __html: previewHtml }}
                  />
                </div>
              ) : (
                <div style={s.previewPlaceholder}>
                  <img src="/placeholder.svg?height=64&width=64" alt="" style={{ opacity: 0.3, marginBottom: 8 }} />
                  <span style={s.previewPlaceholderText}>Preview will appear here</span>
                </div>
              )}
            </div>
            {previewHtml && (
              <div style={{ textAlign: 'right', marginTop: 8 }}>
                <button onClick={handleDownload} style={s.button}>
                  💾 Save As (.pptx)
                </button>
              </div>
            )}
          </div>
        </div>

        {/* Status Bar */}
        <div style={s.statusBar}>
          <div style={s.statusSegment}>
            {status === 'loading' ? '⏳ Generating slide…' : status === 'success' ? '✔ Slide generated successfully' : 'Ready'}
          </div>
          <div style={s.statusSegment}>SNV Corp © 2000</div>
        </div>
      </div>
    </main>
  );
}

/* ─── Windows 2000 style sheet ─── */
const s = {
  desktop: {
    minHeight: '100vh',
    background: 'repeating-linear-gradient(135deg,#008080 0px,#008080 2px,#007070 2px,#007070 4px)',
    display: 'flex',
    alignItems: 'flex-start',
    justifyContent: 'center',
    padding: '32px 16px',
    fontFamily: '"Tahoma","MS Sans Serif",Arial,sans-serif',
    fontSize: 12,
  },
  window: {
    width: '100%',
    maxWidth: 900,
    background: '#d4d0c8',
    border: '2px solid',
    borderColor: '#ffffff #808080 #808080 #ffffff',
    boxShadow: '2px 2px 0 #000',
    display: 'flex',
    flexDirection: 'column',
  },
  titleBar: {
    background: 'linear-gradient(to right,#0a246a,#a6caf0)',
    padding: '3px 4px',
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    userSelect: 'none',
  },
  titleBarLeft: { display: 'flex', alignItems: 'center', gap: 4 },
  titleIcon: { fontSize: 14 },
  titleText: { color: '#fff', fontWeight: 'bold', fontSize: 12 },
  titleBarButtons: { display: 'flex', gap: 2 },
  winBtn: {
    width: 18,
    height: 16,
    fontSize: 10,
    lineHeight: '14px',
    textAlign: 'center',
    padding: 0,
    cursor: 'pointer',
    background: '#d4d0c8',
    border: '1px solid',
    borderColor: '#ffffff #808080 #808080 #ffffff',
    boxShadow: 'inset 1px 1px #dfdfdf',
    color: '#000',
    fontFamily: 'inherit',
  },
  winBtnClose: { background: '#d4d0c8' },
  menuBar: {
    background: '#d4d0c8',
    borderBottom: '1px solid #808080',
    padding: '2px 4px',
    display: 'flex',
    gap: 0,
  },
  menuItem: {
    padding: '2px 8px',
    cursor: 'default',
    fontSize: 12,
    color: '#000',
  },
  windowBody: {
    display: 'flex',
    gap: 8,
    padding: 12,
    flexWrap: 'wrap',
  },
  panel: {
    flex: '1 1 300px',
    display: 'flex',
    flexDirection: 'column',
    gap: 8,
    minWidth: 260,
  },
  groupBox: {
    border: '1px solid',
    borderColor: '#808080 #dfdfdf #dfdfdf #808080',
    padding: '8px 10px 10px',
    margin: 0,
    background: '#d4d0c8',
  },
  legend: {
    fontWeight: 'bold',
    fontSize: 11,
    color: '#000',
    padding: '0 4px',
  },
  label: {
    display: 'block',
    fontSize: 11,
    color: '#000',
    marginBottom: 3,
  },
  input: {
    width: '100%',
    padding: '2px 4px',
    fontSize: 12,
    fontFamily: '"Tahoma","MS Sans Serif",Arial,sans-serif',
    background: '#fff',
    border: '1px solid',
    borderColor: '#808080 #dfdfdf #dfdfdf #808080',
    boxSizing: 'border-box',
    outline: 'none',
    color: '#000',
  },
  fileInput: {
    fontSize: 11,
    fontFamily: '"Tahoma","MS Sans Serif",Arial,sans-serif',
    color: '#000',
    width: '100%',
  },
  fileNote: {
    margin: '4px 0 0',
    fontSize: 11,
    color: '#000080',
  },
  button: {
    padding: '3px 16px',
    fontSize: 12,
    fontFamily: '"Tahoma","MS Sans Serif",Arial,sans-serif',
    background: '#d4d0c8',
    color: '#000',
    border: '2px solid',
    borderColor: '#ffffff #808080 #808080 #ffffff',
    boxShadow: 'inset 1px 1px #dfdfdf',
    cursor: 'pointer',
    minWidth: 75,
  },
  buttonRow: {
    display: 'flex',
    justifyContent: 'flex-end',
    gap: 6,
    marginTop: 4,
  },
  errorBox: {
    background: '#fff8e1',
    border: '1px solid #e0b000',
    color: '#7a4f00',
    padding: '4px 8px',
    fontSize: 11,
    fontFamily: '"Tahoma","MS Sans Serif",Arial,sans-serif',
  },
  previewPanel: {
    flex: '1 1 340px',
    display: 'flex',
    flexDirection: 'column',
    minWidth: 300,
  },
  previewInset: {
    flex: 1,
    minHeight: 330,
    background: '#fff',
    border: '2px solid',
    borderColor: '#808080 #dfdfdf #dfdfdf #808080',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    overflow: 'hidden',
    position: 'relative',
  },
  previewPlaceholder: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    justifyContent: 'center',
    color: '#808080',
    gap: 4,
  },
  previewPlaceholderText: {
    fontSize: 11,
    fontFamily: '"Tahoma","MS Sans Serif",Arial,sans-serif',
    color: '#808080',
  },
  statusBar: {
    background: '#d4d0c8',
    borderTop: '1px solid #808080',
    display: 'flex',
    justifyContent: 'space-between',
    padding: '2px 4px',
  },
  statusSegment: {
    border: '1px solid',
    borderColor: '#808080 #dfdfdf #dfdfdf #808080',
    padding: '1px 6px',
    fontSize: 11,
    color: '#000',
    minWidth: 80,
  },
};
