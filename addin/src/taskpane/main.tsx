import { createRoot } from 'react-dom/client';
import { App } from './App.js';
import './styles.css';
// Shared runtime: the task pane page hosts the custom functions too.
import '../functions/functions.js';

/* global Office */
Office.onReady(() => {
  const container = document.getElementById('root');
  if (container) createRoot(container).render(<App />);
});
