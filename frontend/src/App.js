import React from 'react';
import { BrowserRouter as Router, Routes, Route } from 'react-router-dom';
import DetalhesPedidoPage from './pages/DetalhesPedidoPage';
import './styles/App.css';

function App() {
  return (
    <Router>
      <Routes>
        <Route path="/detalhes-pedido/:numeroPedido" element={<DetalhesPedidoPage />} />
      </Routes>
    </Router>
  );
}

export default App;
