import React, { useEffect, useState } from 'react';
import ReactDOM from 'react-dom/client';
import { FluentProvider } from '@fluentui/react-components';
import { App } from './App';
import { kbDarkTheme, kbLightTheme } from './theme';
import './index.css';

const THEME_STORAGE_KEY = 'kbloader.theme';

function Root() {
  const [themeMode, setThemeMode] = useState<'light' | 'dark'>(() => {
    if (typeof window === 'undefined') {
      return 'light';
    }
    return window.localStorage.getItem(THEME_STORAGE_KEY) === 'dark' ? 'dark' : 'light';
  });

  useEffect(() => {
    window.localStorage.setItem(THEME_STORAGE_KEY, themeMode);
  }, [themeMode]);

  return (
    <FluentProvider theme={themeMode === 'dark' ? kbDarkTheme : kbLightTheme}>
      <App
        themeMode={themeMode}
        onToggleTheme={() => setThemeMode(current => current === 'dark' ? 'light' : 'dark')}
      />
    </FluentProvider>
  );
}

ReactDOM.createRoot(document.getElementById('root')!).render(
  <React.StrictMode>
    <Root />
  </React.StrictMode>
);
