import React, { useEffect, useState } from 'react';
import { getApiBase } from '../../utils/runtimeConfig';

export const NotificationEmailsTab: React.FC = () => {
  const [emails, setEmails] = useState<string[]>([]);
  const [input, setInput] = useState('');
  const [loading, setLoading] = useState(false);
  const [saving, setSaving] = useState(false);
  const [status, setStatus] = useState<string | null>(null);
  const apiBase = (getApiBase() as string) || '';

  const loadEmails = async () => {
    setLoading(true);
    setStatus(null);
    try {
      const res = await fetch(`${apiBase}/api/notification-emails`);
      const j = await res.json();
      setEmails(Array.isArray(j?.emails) ? j.emails : []);
    } catch (e) {
      setStatus('Failed to load emails');
    } finally {
      setLoading(false);
    }
  };
  useEffect(() => {
    void loadEmails();
  }, []);

  const saveEmails = async (next: string[]) => {
    setSaving(true);
    setStatus(null);
    try {
      const res = await fetch(`${apiBase}/api/notification-emails`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ emails: next })
      });
      if (!res.ok) throw new Error('Save failed');
      setEmails(next);
      setStatus('Saved!');
    } catch (e) {
      setStatus('Failed to save');
    } finally {
      setSaving(false);
    }
  };

  const addEmail = () => {
    const val = input.trim().toLowerCase();
    if (!val || !val.includes('@') || emails.includes(val)) return;
    const next = [...emails, val];
    setEmails(next);
    setInput('');
    void saveEmails(next);
  };
  const removeEmail = (email: string) => {
    const next = emails.filter((e) => e !== email);
    setEmails(next);
    void saveEmails(next);
  };

  return (
    <div className="card" style={{ maxWidth: 480, margin: '0 auto', padding: 24 }}>
      <h3 style={{ margin: '0 0 12px 0', fontSize: 18 }}>Notification Emails</h3>
      <div className="small muted" style={{ marginBottom: 12 }}>
        These emails will receive admin notifications (batch completions, nudges, etc). Changes are saved instantly.
      </div>
      <div style={{ display: 'flex', gap: 8, marginBottom: 16 }}>
        <input
          type="email"
          value={input}
          onChange={(e) => setInput(e.target.value)}
          placeholder="admin@domain.com"
          style={{ flex: 1, padding: 8, border: '1px solid #ddd', borderRadius: 4 }}
          disabled={saving}
        />
        <button
          className="btn sm"
          onClick={addEmail}
          disabled={
            saving || !input.trim() || !input.includes('@') || emails.includes(input.trim().toLowerCase())
          }
        >
          Add
        </button>
      </div>
      {loading ? (
        <div className="small muted">Loading...</div>
      ) : emails.length === 0 ? (
        <div className="small muted">No notification emails set.</div>
      ) : (
        <ul style={{ listStyle: 'none', padding: 0, margin: 0 }}>
          {emails.map((email) => (
            <li
              key={email}
              style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 6 }}
            >
              <span style={{ flex: 1 }}>{email}</span>
              <button className="btn ghost sm" onClick={() => removeEmail(email)} disabled={saving}>
                Remove
              </button>
            </li>
          ))}
        </ul>
      )}
      {status && (
        <div className="small muted" style={{ marginTop: 8 }}>
          {status}
        </div>
      )}
    </div>
  );
};
