import React from 'react';
import './CabecalhoDetalhes.css';

function CabecalhoDetalhes({ pedido }) {
  return (
    <div className="cabecalho-detalhes">
      <h2>Detalhes do Pedido</h2>
      <p><strong>Número do Pedido:</strong> {pedido.Pedido}</p>
      <p><strong>CNPJ:</strong> {pedido.Cliente.CNPJ}</p>
      <p><strong>Razão Social:</strong> {pedido.Cliente.RazaoSocial}</p>
    </div>
  );
}

export default CabecalhoDetalhes;
