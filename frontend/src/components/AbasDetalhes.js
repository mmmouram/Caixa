import React, { useState } from 'react';
import ListaDados from './ListaDados';
import PedidoService from '../services/PedidoService';
import ModalDetalhes from './ModalDetalhes';
import './AbasDetalhes.css';

function AbasDetalhes({ numeroPedido, abas }) {
  const [abaAtiva, setAbaAtiva] = useState('Itens');
  const [dadosAba, setDadosAba] = useState(abas);
  const [ordem, setOrdem] = useState(null);
  const [modalDados, setModalDados] = useState(null);

  const handleTrocaAba = async (aba) => {
    setAbaAtiva(aba);
    try {
      const dados = await PedidoService.obterDadosAba(numeroPedido, aba, ordem);
      setDadosAba({ [aba]: dados });
    } catch (error) {
      alert(error.message || "Erro ao carregar dados da aba.");
    }
  };

  const handleAtualizar = async () => {
    try {
      const dados = await PedidoService.obterDadosAba(numeroPedido, abaAtiva, ordem);
      setDadosAba({ [abaAtiva]: dados });
      alert("Dados atualizados com sucesso.");
    } catch (error) {
      alert(error.message || "Erro ao atualizar os dados.");
    }
  };

  const handleExportar = async () => {
    try {
      await PedidoService.exportarExcel(numeroPedido, abaAtiva);
      alert("Exportação realizada com sucesso.");
    } catch (error) {
      alert(error.message || "Erro ao exportar os dados.");
    }
  };

  const handleOrdenar = async (coluna) => {
    const novaOrdem = ordem === coluna ? `${coluna}_desc` : coluna;
    setOrdem(novaOrdem);
    try {
      const dados = await PedidoService.obterDadosAba(numeroPedido, abaAtiva, novaOrdem);
      setDadosAba({ [abaAtiva]: dados });
    } catch (error) {
      alert(error.message || "Erro ao ordenar os dados.");
    }
  };

  const handleDoubleClick = (item) => {
    // Ao clicar duas vezes, abrir modal com detalhes do item
    setModalDados(item);
  };

  // Definindo as abas padrão do sistema
  const abasNomes = ['Itens', 'Observacoes', 'Bloqueios', 'NotasFiscais'];

  return (
    <div className="abas-detalhes">
      <div className="abas-menu">
        {abasNomes.map((aba) => (
          <button
            key={aba}
            className={abaAtiva === aba ? 'ativa' : ''}
            onClick={() => handleTrocaAba(aba)}
          >
            {aba}
          </button>
        ))}
      </div>
      <div className="acoes">
        <button onClick={handleAtualizar}>Atualizar</button>
        <button onClick={handleExportar}>Exportar Excel</button>
      </div>
      <div className="conteudo-aba">
        <ListaDados
          dados={dadosAba[abaAtiva] || []}
          onOrdenar={handleOrdenar}
          onDoubleClick={handleDoubleClick}
        />
      </div>
      {modalDados && (
        <ModalDetalhes dados={modalDados} onClose={() => setModalDados(null)} />
      )}
    </div>
  );
}

export default AbasDetalhes;
