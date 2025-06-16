const API_BASE = "/api/DetalhesPedido";

const PedidoService = {
  obterDetalhesPedido: async (numeroPedido) => {
    const response = await fetch(`${API_BASE}/${numeroPedido}`);
    if (!response.ok) {
      throw new Error("Pedido não encontrado.");
    }
    return await response.json();
  },
  obterDadosAba: async (numeroPedido, aba, ordenacao = null) => {
    let url = `${API_BASE}/${numeroPedido}/aba/${aba}`;
    if (ordenacao) {
      url += `?ordenacao=${ordenacao}`;
    }
    const response = await fetch(url);
    if (!response.ok) {
      throw new Error("Erro ao buscar dados da aba.");
    }
    return await response.json();
  },
  exportarExcel: async (numeroPedido, aba) => {
    const response = await fetch(`${API_BASE}/${numeroPedido}/aba/${aba}/exportar`, { method: 'POST' });
    if (!response.ok) {
      throw new Error("Erro ao exportar os dados.");
    }
    // Realiza o download do arquivo Excel
    const blob = await response.blob();
    const url = window.URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `${aba}_pedido_${numeroPedido}.xlsx`;
    document.body.appendChild(link);
    link.click();
    link.remove();
    return true;
  }
};

export default PedidoService;
