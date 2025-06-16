import React from 'react';
import { render, screen, waitFor } from '@testing-library/react';
import { MemoryRouter, Route, Routes } from 'react-router-dom';
import DetalhesPedidoPage from './DetalhesPedidoPage';
import PedidoService from '../services/PedidoService';

jest.mock('../services/PedidoService');

const mockPedido = {
  Pedido: '123',
  Cliente: { CNPJ: '00.000.000/0001-00', RazaoSocial: 'Cliente Teste' },
  Abas: { 
    Itens: [{ Item: 'Produto1' }], 
    Observacoes: [], 
    Bloqueios: [], 
    NotasFiscais: []
  }
};

describe('DetalhesPedidoPage', () => {
  test('renders pedido details on successful fetch', async () => {
    PedidoService.obterDetalhesPedido.mockResolvedValue(mockPedido);
    render(
      <MemoryRouter initialEntries={['/detalhes-pedido/123']}>
        <Routes>
          <Route path="/detalhes-pedido/:numeroPedido" element={<DetalhesPedidoPage />} />
        </Routes>
      </MemoryRouter>
    );

    // Initially shows loading state
    expect(screen.getByText(/Carregando.../i)).toBeInTheDocument();

    // Wait for the async fetch to complete and the page update
    await waitFor(() => {
      expect(screen.getByText(/Número do Pedido:/i)).toBeInTheDocument();
      expect(screen.getByText(/00.000.000\/0001-00/i)).toBeInTheDocument();
      expect(screen.getByText(/Cliente Teste/i)).toBeInTheDocument();
    });
  });

  test('renders error message on fetch failure', async () => {
    PedidoService.obterDetalhesPedido.mockRejectedValue(new Error('Pedido não encontrado.'));
    render(
      <MemoryRouter initialEntries={['/detalhes-pedido/999']}>
        <Routes>
          <Route path="/detalhes-pedido/:numeroPedido" element={<DetalhesPedidoPage />} />
        </Routes>
      </MemoryRouter>
    );
    
    await waitFor(() => {
      expect(screen.getByText(/Pedido não encontrado./i)).toBeInTheDocument();
    });
  });
});
