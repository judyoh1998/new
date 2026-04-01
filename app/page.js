'use client';

import { useState } from 'react';
import { supabase } from '../lib/supabase';

export default function Home() {
  const [form, setForm] = useState({ title: '', description: '', content: '' });
  const [status, setStatus] = useState(null); // 'loading' | 'success' | 'error'

  function handleChange(e) {
    setForm({ ...form, [e.target.name]: e.target.value });
  }

  async function handleGenerate() {
    if (!form.title.trim() || !form.description.trim() || !form.content.trim()) {
      setStatus('validation');
      return;
    }
    setStatus('loading');
    const { error } = await supabase.from('slides').insert([
      {
        title: form.title,
        description: form.description,
        content: form.content,
      },
    ]);
    if (error) {
      setStatus('error');
    } else {
      setStatus('success');
      setForm({ title: '', description: '', content: '' });
    }
  }

  return (
    <main style={styles.main}>
      <div style={styles.card}>
        <h1 style={styles.heading}>Slide Generator</h1>

        <label style={styles.label}>Slide Title</label>
        <input
          name="title"
          value={form.title}
          onChange={handleChange}
          placeholder="Enter slide title"
          style={styles.input}
        />

        <label style={styles.label}>Slide Description</label>
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

        {status === 'validation' && (
          <p style={styles.error}>Please fill in all fields before generating.</p>
        )}
        {status === 'success' && (
          <p style={styles.success}>Slide saved successfully!</p>
        )}
        {status === 'error' && (
          <p style={styles.error}>Something went wrong. Please try again.</p>
        )}
      </div>
    </main>
  );
}

const styles = {
  main: {
    minHeight: '100vh',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    padding: '24px',
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
  success: {
    color: '#2e7d32',
    fontSize: '14px',
    marginTop: '8px',
  },
  error: {
    color: '#c62828',
    fontSize: '14px',
    marginTop: '8px',
  },
};
