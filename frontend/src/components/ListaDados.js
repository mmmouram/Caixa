import React from 'react';
import './ListaDados.css';

function ListaDados({ dados, onOrdenar, onDoubleClick }) {
  if (!dados || dados.length === 0) {
    return <p className="centralizar">Nenhum dado encontrado.</p>;
  }

  const colunas = Object.keys(dados[0]);

  return (
    <table className="lista-dados">
      <thead>
        <tr>
          {colunas.map((coluna) => (
            <th key={coluna} onClick={() => onOrdenar(coluna)}>
              {coluna}
            </th>
          ))}
        </tr>
      </thead>
      <tbody>
        {dados.map((item, index) => (
          <tr key={index} onDoubleClick={() => onDoubleClick(item)}>
            {colunas.map((coluna) => (
              <td key={coluna}>{item[coluna]}</td>
            ))}
          </tr>
        ))}
      </tbody>
    </table>
  );
}

export default ListaDados;
