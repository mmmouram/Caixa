import React from 'react';
import { render, screen, fireEvent, waitFor } from '@testing-library/react';
import AbasDetalhes from './AbasDetalhes';
import PedidoService from '../services/PedidoService';

jest.mock('../services/PedidoService');

const initialAbas = { 
  Itens: [{ Nome: 'Produto1', Valor: '10' }],
  Observacoes: [],
  Bloqueios: [],
  NotasFiscais: []
};

describe('AbasDetalhes component', () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  test('renders abas and lists data for default active tab', async () => {
    PedidoService.obterDadosAba.mockResolvedValue(initialAbas.Itens);
    render(<AbasDetalhes numeroPedido="123" abas={initialAbas} />);

    // Check that default active tab is "Itens"
    expect(screen.getByText('Itens')).toHaveClass('ativa');

    // Check that data from the Itens tab is rendered
    expect(await screen.findByText('Produto1')).toBeInTheDocument();
  });

  test('calls handleAtualizar and shows alert on success', async () => {
    window.alert = jest.fn();
    PedidoService.obterDadosAba.mockResolvedValue(initialAbas.Itens);
    render(<AbasDetalhes numeroPedido="123" abas={initialAbas} />);
    
    const atualizarButton = screen.getByText('Atualizar');
    fireEvent.click(atualizarButton);

    await waitFor(() => {
      expect(PedidoService.obterDadosAba).toHaveBeenCalled();
      expect(window.alert).toHaveBeenCalledWith('Dados atualizados com sucesso.');
    });
  });

  test('calls handleExportar and shows alert on success', async () => {
    window.alert = jest.fn();
    PedidoService.exportarExcel.mockResolvedValue(true);
    render(<AbasDetalhes numeroPedido="123" abas={initialAbas} />);
    
    const exportarButton = screen.getByText('Exportar Excel');
    fireEvent.click(exportarButton);

    await waitFor(() => {
      expect(PedidoService.exportarExcel).toHaveBeenCalled();
      expect(window.alert).toHaveBeenCalledWith('Exportação realizada com sucesso.');
    });
  });
});
