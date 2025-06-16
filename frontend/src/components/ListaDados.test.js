import React from 'react';
import { render, screen, fireEvent } from '@testing-library/react';
import ListaDados from './ListaDados';

describe('ListaDados component', () => {
  const dados = [
    { Nome: 'Produto1', Valor: '10' },
    { Nome: 'Produto2', Valor: '20' }
  ];

  test('renders table with data', () => {
    const mockOrdenar = jest.fn();
    const mockDoubleClick = jest.fn();
    render(<ListaDados dados={dados} onOrdenar={mockOrdenar} onDoubleClick={mockDoubleClick} />);

    // Check that headers are rendered
    expect(screen.getByText('Nome')).toBeInTheDocument();
    expect(screen.getByText('Valor')).toBeInTheDocument();
    
    // Check that row data is rendered
    expect(screen.getByText('Produto1')).toBeInTheDocument();
    expect(screen.getByText('Produto2')).toBeInTheDocument();
  });

  test('calls onOrdenar when header is clicked', () => {
    const mockOrdenar = jest.fn();
    const mockDoubleClick = jest.fn();
    render(<ListaDados dados={dados} onOrdenar={mockOrdenar} onDoubleClick={mockDoubleClick} />);
    
    const header = screen.getByText('Nome');
    fireEvent.click(header);
    expect(mockOrdenar).toHaveBeenCalledWith('Nome');
  });

  test('calls onDoubleClick when a row is double-clicked', () => {
    const mockOrdenar = jest.fn();
    const mockDoubleClick = jest.fn();
    render(<ListaDados dados={dados} onOrdenar={mockOrdenar} onDoubleClick={mockDoubleClick} />);
    
    const row = screen.getByText('Produto1').closest('tr');
    fireEvent.doubleClick(row);
    expect(mockDoubleClick).toHaveBeenCalledWith(expect.objectContaining({ Nome: 'Produto1', Valor: '10' }));
  });
});
