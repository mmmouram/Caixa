import React from 'react';
import { render } from '@testing-library/react';
import App from './App';

// Snapshot test to ensure App renders correctly

test('App renders without crashing', () => {
  const { container } = render(<App />);
  expect(container).toBeDefined();
  expect(container).toMatchSnapshot();
});
