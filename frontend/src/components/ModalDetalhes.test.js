import React from 'react';
import { render, screen, fireEvent } from '@testing-library/react';
import ModalDetalhes from './ModalDetalhes';


test('renders modal with details and calls onClose', () => {
  const dados = { Campo1: 'Valor1', Campo2: 'Valor2' };
  const onClose = jest.fn();
  render(<ModalDetalhes dados={dados} onClose={onClose} />);
  
  // Check header and detail fields
  expect(screen.getByText('Detalhes do Item')).toBeInTheDocument();
  expect(screen.getByText('Campo1:')).toBeInTheDocument();
  expect(screen.getByText('Valor1')).toBeInTheDocument();
  
  const closeButton = screen.getByText('Fechar');
  fireEvent.click(closeButton);
  expect(onClose).toHaveBeenCalled();
});
