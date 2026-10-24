interface ErrorDialogProps {
  show: boolean;
  message: string;
  onClose: () => void;
  variant?: 'error' | 'success';
}

export const ErrorDialog = ({ show, message, onClose, variant = 'error' }: ErrorDialogProps) => {
  if (!show) return null;
  const isSuccess = variant === 'success';

  return (
    <div style={{
      position: 'fixed',
      top: 0,
      left: 0,
      right: 0,
      bottom: 0,
      backgroundColor: 'rgba(0, 0, 0, 0.5)',
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'center',
      zIndex: 2400
    }} onClick={onClose}>
      <div style={{
        backgroundColor: 'white',
        padding: '2rem',
        borderRadius: '8px',
        maxWidth: '600px',
        width: '90%',
        maxHeight: '80vh',
        overflow: 'auto',
        boxShadow: '0 4px 20px rgba(0,0,0,0.3)'
      }} onClick={(e) => e.stopPropagation()}>
        <h2 style={{ 
          margin: '0 0 1rem 0', 
          color: isSuccess ? '#2e7d32' : '#d32f2f',
          display: 'flex',
          alignItems: 'center',
          gap: '10px'
        }}>
          <span style={{ fontSize: '2rem' }}>{isSuccess ? '✓' : '⚠️'}</span>
          {isSuccess ? 'Saved' : 'Error'}
        </h2>
        <pre style={{
          backgroundColor: isSuccess ? '#e8f5e9' : '#ffebee',
          padding: '1rem',
          borderRadius: '4px',
          color: isSuccess ? '#1b5e20' : '#c62828',
          whiteSpace: 'pre-wrap',
          wordBreak: 'break-word',
          fontSize: '14px',
          lineHeight: '1.6',
          margin: '0 0 1.5rem 0'
        }}>
          {message}
        </pre>
        <button 
          onClick={onClose}
          style={{
            backgroundColor: isSuccess ? '#2e7d32' : '#d32f2f',
            color: 'white',
            border: 'none',
            padding: '10px 24px',
            borderRadius: '4px',
            cursor: 'pointer',
            fontSize: '16px',
            fontWeight: 'bold'
          }}
        >
          Close
        </button>
      </div>
    </div>
  );
};
