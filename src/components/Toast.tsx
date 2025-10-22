import React from 'react';

type ToastProps = { message: string; show: boolean };

const Toast: React.FC<ToastProps> = ({ message, show }) => {
  return (
    <div className={`toast${show ? ' show' : ''}`} role="status" aria-live="polite">
      {message}
    </div>
  );
};

export default Toast;
