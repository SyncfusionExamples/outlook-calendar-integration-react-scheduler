import React from 'react';
import logo from './logo.svg';
import './App.css';
import { IPublicClientApplication } from '@azure/msal-browser';
import { MsalProvider } from '@azure/msal-react';
import { BrowserRouter as Router, Routes, Route } from 'react-router-dom';
import ProvideAppContext from './AppContext';
import { Container } from 'react-bootstrap';
import Scheduler from './Scheduler';
import NavBar from './NavBar';
import Welcome from './Welcome';
import ErrorMessage from './ErrorMessage';
import 'bootstrap/dist/css/bootstrap.css';

type AppProps = {
  pca: IPublicClientApplication
};

function App({ pca }: AppProps) {
  return (
    <MsalProvider instance={pca}>
      <ProvideAppContext>
        <Router>
          <NavBar />
          <Container>
            <ErrorMessage />
            <Routes>
              <Route path="/"
                element={
                  <Welcome />
                } />
              <Route path="/scheduler"
                element={
                  <Scheduler />
                } />              
            </Routes>
          </Container>
        </Router>
      </ProvideAppContext>
    </MsalProvider>
  );
}

export default App;
