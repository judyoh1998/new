'use client';

import { useState } from 'react';

export default function Home() {
  const [brandFile, setBrandFile] = useState(null);
  const [form, setForm] = useState({ title: '', description: '', content: '' });
  const [status, setStatus] = useState(null); // 'loading' | 'success' | 'error' | 'validation'
  const [errorMsg, setErrorMsg] = useState(null);
  const [previewHtml, setPreviewHtml] = useState(null);
  const [pptxBase64, setPptxBase64] = useState(null);

  function handleChange(e) {
    setForm({ ...form, [e.target.name]: e.target.value });
  }

  async function handleGenerate() {
    if (!brandFile) {
      setStatus('no-file');
      return;
    }
    if (!form.title.trim() || !form.description.trim() || !form.content.trim()) {
      setStatus('validation');
      return;
    }

    setStatus('loading');
    setPreviewHtml(null);
    setPptxBase64(null);
    setErrorMsg(null);

    // Read reference image as base64
    let imageBase64 = null;
    let imageMimeType = null;
    try {
      const dataUrl = await new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => resolve(e.target.result);
        reader.onerror = reject;
        reader.readAsDataURL(brandFile);
      });
      const [header, b64] = dataUrl.split(',');
      imageBase64 = b64;
      imageMimeType = header.match(/:(.*?);/)[1];
    } catch {
      // Continue without reference image if reading fails
    }

    try {
      const res = await fetch('/api/generate', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ title: form.title, description: form.description, content: form.content, imageBase64, imageMimeType }),
      });
      const data = await res.json();
      if (!res.ok) {
        setErrorMsg(data?.error || 'Unknown error');
        setStatus('error');
        return;
      }
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
    const blob = new Blob([bytes], {
      type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${form.title || 'slide'}.pptx`;
    a.click();
    URL.revokeObjectURL(url);
  }

  return (
    <main style={styles.main}>
      <div style={styles.card}>
        <h1 style={styles.heading}>SNV Slide Generator</h1>

        {/* Brand file upload */}
        <label style={styles.label}>Reference Slide Image (PNG or JPG)</label>
        <input
          type="file"
          accept="image/png,image/jpeg,image/jpg"
          onChange={(e) => setBrandFile(e.target.files?.[0] || null)}
          style={styles.fileInput}
        />
        {brandFile && (
          <p style={styles.fileNote}>{brandFile.name}</p>
        )}

        <div style={styles.divider} />

        {/* Slide form */}
        <label style={styles.label}>Slide Title</label>
        <input
          name="title"
          value={form.title}
          onChange={handleChange}
          placeholder="Enter slide title"
          style={styles.input}
        />

        <label style={styles.label}>Short Description of Your Slide!</label>
        <input
          name="description"
          value={form.description}
          onChange={handleChange}
          placeholder="Enter slide description"
          style={styles.input}
        />

        <label style={styles.label}>Slide Content</label>
        <textarea
          name="content"
          value={form.content}
          onChange={handleChange}
          placeholder="Enter slide content"
          rows={5}
          style={{ ...styles.input, resize: 'vertical' }}
        />

        <button
          onClick={handleGenerate}
          disabled={status === 'loading'}
          style={{
            ...styles.button,
            opacity: status === 'loading' ? 0.6 : 1,
            cursor: status === 'loading' ? 'not-allowed' : 'pointer',
          }}
        >
          {status === 'loading' ? 'Generating...' : 'Generate'}
        </button>

        {status === 'no-file' && <p style={styles.error}>Please upload a reference slide image first.</p>}
        {status === 'validation' && <p style={styles.error}>Please fill in all fields before generating.</p>}
        {status === 'error' && (
          <p style={styles.error}>{errorMsg || 'Something went wrong. Please try again.'}</p>
        )}
      </div>

      {/* Preview */}
      {previewHtml && (
        <div style={styles.previewCard}>
          <div style={styles.previewHeader}>
            <span style={styles.previewLabel}>Preview</span>
            <button onClick={handleDownload} style={styles.downloadButton}>
              Download .pptx
            </button>
          </div>
          <div style={{ overflowX: 'auto' }}>
            <div
              style={{ transform: 'scale(0.75)', transformOrigin: 'top left', width: 800, height: 450 }}
              dangerouslySetInnerHTML={{ __html: previewHtml }}
            />
          </div>
        </div>
      )}
    </main>
  );
}

const styles = {
  main: {
    minHeight: '100vh',
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    padding: '40px 24px',
    gap: '24px',
  },
  card: {
    background: '#fff',
    borderRadius: '12px',
    padding: '40px',
    width: '100%',
    maxWidth: '560px',
    boxShadow: '0 2px 12px rgba(0,0,0,0.08)',
    display: 'flex',
    flexDirection: 'column',
    gap: '6px',
  },
  heading: {
    margin: '0 0 24px',
    fontSize: '22px',
    fontWeight: '600',
    color: '#111',
  },
  label: {
    fontSize: '14px',
    fontWeight: '500',
    color: '#444',
    marginTop: '12px',
  },
  input: {
    width: '100%',
    padding: '10px 12px',
    fontSize: '15px',
    border: '1px solid #ddd',
    borderRadius: '8px',
    outline: 'none',
    boxSizing: 'border-box',
    marginTop: '4px',
  },
  fileInput: {
    marginTop: '6px',
    fontSize: '14px',
    color: '#444',
  },
  fileNote: {
    margin: '4px 0 0',
    fontSize: '13px',
    color: '#666',
  },
  divider: {
    height: '1px',
    background: '#eee',
    margin: '16px 0 4px',
  },
  button: {
    marginTop: '24px',
    padding: '12px',
    fontSize: '15px',
    fontWeight: '600',
    background: '#111',
    color: '#fff',
    border: 'none',
    borderRadius: '8px',
    width: '100%',
  },
  error: {
    color: '#c62828',
    fontSize: '14px',
    marginTop: '8px',
  },
  previewCard: {
    background: '#fff',
    borderRadius: '12px',
    padding: '24px',
    width: '100%',
    maxWidth: '640px',
    boxShadow: '0 2px 12px rgba(0,0,0,0.08)',
  },
  previewHeader: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    marginBottom: '16px',
  },
  previewLabel: {
    fontSize: '14px',
    fontWeight: '600',
    color: '#444',
  },
  downloadButton: {
    padding: '8px 16px',
    fontSize: '13px',
    fontWeight: '600',
    background: '#111',
    color: '#fff',
    border: 'none',
    borderRadius: '6px',
    cursor: 'pointer',
  },
};
