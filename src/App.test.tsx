import React from 'react';
import { render, screen } from '@testing-library/react';
import App from './App';
import { PublicClientApplication } from '@azure/msal-browser';

let pca: PublicClientApplication;

test('renders learn react link', () => {
 
  render(<App pca={pca}/>);
  const linkElement = screen.getByText(/learn react/i);
  expect(linkElement).toBeInTheDocument();
});
