import PedidoService from './PedidoService';

describe('PedidoService', () => {
  beforeEach(() => {
    // Optionally reset fetch mock if using jest-fetch-mock
    if (fetch.resetMocks) {
      fetch.resetMocks();
    }
  });

  test('obterDetalhesPedido returns data on success', async () => {
    const mockData = { Pedido: '123' };
    global.fetch = jest.fn().mockResolvedValue({
      ok: true,
      json: async () => mockData
    });
    
    const data = await PedidoService.obterDetalhesPedido('123');
    expect(data).toEqual(mockData);
    expect(fetch).toHaveBeenCalledWith('/api/DetalhesPedido/123');
  });

  test('obterDetalhesPedido throws error on failure', async () => {
    global.fetch = jest.fn().mockResolvedValue({ ok: false });
    await expect(PedidoService.obterDetalhesPedido('999')).rejects.toThrow('Pedido não encontrado.');
  });

  test('obterDadosAba returns data on success', async () => {
    const mockData = [{ item: '1' }];
    global.fetch = jest.fn().mockResolvedValue({
      ok: true,
      json: async () => mockData
    });
    
    const data = await PedidoService.obterDadosAba('123', 'Itens');
    expect(data).toEqual(mockData);
    expect(fetch).toHaveBeenCalledWith('/api/DetalhesPedido/123/aba/Itens');
  });

  test('exportarExcel triggers download and returns true on success', async () => {
    // Create a fake blob
    const fakeBlob = new Blob(['dummy content'], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    global.fetch = jest.fn().mockResolvedValue({
      ok: true,
      blob: async () => fakeBlob
    });
    
    // Prepare DOM for download link
    document.body.innerHTML = '';
    const urlSpy = jest.spyOn(window.URL, 'createObjectURL').mockReturnValue('blob:http://dummy');
    const appendChildSpy = jest.spyOn(document.body, 'appendChild');
    
    const result = await PedidoService.exportarExcel('123', 'Itens');
    expect(result).toBe(true);
    expect(fetch).toHaveBeenCalledWith('/api/DetalhesPedido/123/aba/Itens/exportar', { method: 'POST' });

    urlSpy.mockRestore();
    appendChildSpy.mockRestore();
  });

  test('exportarExcel throws error on failure', async () => {
    global.fetch = jest.fn().mockResolvedValue({ ok: false });
    await expect(PedidoService.exportarExcel('123', 'Itens')).rejects.toThrow('Erro ao exportar os dados.');
  });
});
