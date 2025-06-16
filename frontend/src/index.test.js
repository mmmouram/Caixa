import React from 'react';
import { createRoot } from 'react-dom/client';
import App from './App';

test('renders without crashing', () => {
  const rootElement = document.createElement('div');
  rootElement.setAttribute('id', 'root');
  document.body.appendChild(rootElement);
  const root = createRoot(rootElement);
  root.render(<App />);
});
