import React from 'react';
import ReactDOM from 'react-dom/client';
import { FluentProvider } from '@fluentui/react-components';
import { App } from './App';
import { kbLightTheme } from './theme';
import './index.css';

ReactDOM.createRoot(document.getElementById('root')!).render(
  <React.StrictMode>
    <FluentProvider theme={kbLightTheme}>
      <App />
    </FluentProvider>
  </React.StrictMode>
);
