import React from 'react';
import ReactDOM from 'react-dom/client';
import App from './DashboardApp';
import './styles.css';
import './dashboard.css';

ReactDOM.createRoot(document.getElementById('root') as HTMLElement).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>,
);
