import React from 'react';
import { render, screen } from '@testing-library/react';
import CabecalhoDetalhes from './CabecalhoDetalhes';

const mockPedido = {
  Pedido: '123',
  Cliente: { CNPJ: '00.000.000/0001-00', RazaoSocial: 'Cliente Teste' }
};

test('renders pedido header details', () => {
  render(<CabecalhoDetalhes pedido={mockPedido} />);
  
  expect(screen.getByText(/Detalhes do Pedido/i)).toBeInTheDocument();
  expect(screen.getByText(/Número do Pedido:/i)).toBeInTheDocument();
  expect(screen.getByText(/123/i)).toBeInTheDocument();
  expect(screen.getByText(/00.000.000\/0001-00/i)).toBeInTheDocument();
  expect(screen.getByText(/Cliente Teste/i)).toBeInTheDocument();
});
