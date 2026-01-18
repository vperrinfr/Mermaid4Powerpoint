/// <reference types="office-js" />

import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';
import App from './App';

Office.onReady(() => {
  const rootElement = document.getElementById('root');
  if (rootElement) {
    ReactDOM.render(
      <FluentProvider theme={webLightTheme}>
        <App />
      </FluentProvider>,
      rootElement
    );
  }
});
