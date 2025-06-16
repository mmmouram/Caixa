import React, { useState, useEffect } from 'react';
import { useParams } from 'react-router-dom';
import PedidoService from '../services/PedidoService';
import CabecalhoDetalhes from '../components/CabecalhoDetalhes';
import AbasDetalhes from '../components/AbasDetalhes';

function DetalhesPedidoPage() {
  const { numeroPedido } = useParams();
  const [pedido, setPedido] = useState(null);
  const [erro, setErro] = useState(null);

  useEffect(() => {
    async function buscarDetalhes() {
      try {
        const dados = await PedidoService.obterDetalhesPedido(numeroPedido);
        setPedido(dados);
      } catch (error) {
        setErro(error.message || "Erro ao buscar dados do pedido.");
      }
    }
    buscarDetalhes();
  }, [numeroPedido]);

  if (erro) {
    return (
      <div className="centralizar">
        <p>{erro}</p>
      </div>
    );
  }

  if (!pedido) {
    return (
      <div className="centralizar">
        <p>Carregando...</p>
      </div>
    );
  }

  return (
    <div className="detalhes-pedido-container">
      <CabecalhoDetalhes pedido={pedido} />
      <AbasDetalhes numeroPedido={numeroPedido} abas={pedido.Abas} />
    </div>
  );
}

export default DetalhesPedidoPage;
