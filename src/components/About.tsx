import React from 'react';

const About: React.FC = () => (
  <div className="container">
    <div className="card" style={{ maxWidth: 920, margin: '0 auto' }}>
      <div className="title">About this Portal</div>
      <div className="small" style={{ marginTop: 8 }}>A quick overview of how Sunbeth’s Document Acknowledgement works.</div>
      <hr style={{ margin: '14px 0', border: 'none', borderTop: '1px solid #f2f2f2' }} />
      <div style={{ color: '#333', lineHeight: 1.55 }}>
        <p>This portal helps employees read and acknowledge mandatory company documents like corporate policies, health and safety guidance, and periodic updates.</p>
        <ul className="features">
          <li><strong>Simple workflow</strong> — Read each document and acknowledge with a click.</li>
          <li><strong>Progress tracking</strong> — See your overall completion by batch.</li>
          <li><strong>Secure sign-in</strong> — Authentication via Microsoft Entra (Azure AD).</li>
          <li><strong>Compliance</strong> — Acknowledgements are recorded in Dataverse.</li>
        </ul>
        <p>When you’re ready, sign in with your corporate account to view your assigned batches.</p>
      </div>
    </div>
  </div>
);

export default About;
