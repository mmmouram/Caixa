import React from 'react';
import './ModalDetalhes.css';

function ModalDetalhes({ dados, onClose }) {
  return (
    <div className="modal-overlay">
      <div className="modal-conteudo">
        <h3>Detalhes do Item</h3>
        <div className="modal-body">
          {Object.entries(dados).map(([chave, valor]) => (
            <p key={chave}><strong>{chave}:</strong> {valor}</p>
          ))}
        </div>
        <button onClick={onClose}>Fechar</button>
      </div>
    </div>
  );
}

export default ModalDetalhes;
