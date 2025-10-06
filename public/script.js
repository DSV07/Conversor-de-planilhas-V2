// Elementos da interface
const formUpload = document.getElementById('formUpload');
const arquivoInput = document.getElementById('arquivo');
const fileName = document.getElementById('fileName');
const uploadBtn = document.getElementById('uploadBtn');
const uploadStatus = document.getElementById('uploadStatus');
const filtersSection = document.getElementById('filtersSection');
const filtersPlaceholder = document.getElementById('filtersPlaceholder');
const unidadeSelect = document.getElementById('unidadeSelect');
const previewBtn = document.getElementById('previewBtn');
const baixarBtn = document.getElementById('baixarBtn');
const downloadStatus = document.getElementById('downloadStatus');
const previewTable = document.getElementById('preview');
const previewCount = document.getElementById('previewCount');

// Elementos para múltiplos arquivos
const tabs = document.querySelectorAll('.tab');
const tabContents = document.querySelectorAll('.tab-content');
const formUploadMultiple = document.getElementById('formUploadMultiple');
const arquivosInput = document.getElementById('arquivos');
const filesName = document.getElementById('filesName');
const uploadMultipleBtn = document.getElementById('uploadMultipleBtn');
const uploadMultipleStatus = document.getElementById('uploadMultipleStatus');
const arquivosLista = document.getElementById('arquivosLista');
const arquivosCarregados = document.getElementById('arquivosCarregados');
const mergeSection = document.getElementById('mergeSection');
const mergePlaceholder = document.getElementById('mergePlaceholder');
const mesclarBtn = document.getElementById('mesclarBtn');
const mergeStatus = document.getElementById('mergeStatus');

// Estado da aplicação
let fileId = null;
let arquivosMultiplos = [];

// Alternar entre abas
tabs.forEach(tab => {
    tab.addEventListener('click', () => {
        const tabId = tab.getAttribute('data-tab');

        tabs.forEach(t => t.classList.remove('active'));
        tabContents.forEach(tc => tc.classList.remove('active'));

        tab.classList.add('active');
        document.getElementById(`${tabId}-tab`).classList.add('active');
    });
});

// Atualizar nome do arquivo selecionado (single)
arquivoInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
        fileName.textContent = e.target.files[0].name;
    } else {
        fileName.textContent = '';
    }
});

// Atualizar nomes dos arquivos selecionados (multiple)
arquivosInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
        filesName.textContent = `${e.target.files.length} arquivo(s) selecionado(s)`;
    } else {
        filesName.textContent = '';
    }
});

// Permitir arrastar e soltar arquivos
const fileInputLabels = document.querySelectorAll('.file-input-label');

fileInputLabels.forEach(label => {
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        label.addEventListener(eventName, preventDefaults, false);
    });

    ['dragenter', 'dragover'].forEach(eventName => {
        label.addEventListener(eventName, highlight, false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
        label.addEventListener(eventName, unhighlight, false);
    });

    label.addEventListener('drop', handleDrop, false);
});

function preventDefaults(e) {
    e.preventDefault();
    e.stopPropagation();
}

function highlight(e) {
    e.currentTarget.classList.add('highlight');
}

function unhighlight(e) {
    e.currentTarget.classList.remove('highlight');
}

function handleDrop(e) {
    const dt = e.dataTransfer;
    const files = dt.files;
    const input = e.currentTarget.previousElementSibling;

    if (input.multiple) {
        input.files = files;
        if (files.length > 0) {
            filesName.textContent = `${files.length} arquivo(s) selecionado(s)`;
        }
    } else {
        input.files = files;
        if (files.length > 0) {
            fileName.textContent = files[0].name;
        }
    }
}

// Envio do formulário de upload único
formUpload.addEventListener('submit', async e => {
    e.preventDefault();

    if (!arquivoInput.files.length) {
        showStatus(uploadStatus, 'Por favor, selecione um arquivo.', 'error');
        return;
    }

    const arquivo = e.target.arquivo.files[0];
    const formData = new FormData();
    formData.append('arquivo', arquivo);

    uploadBtn.disabled = true;
    uploadBtn.innerHTML = '<div class="spinner"></div><span>Processando...</span>';

    // Mostrar barra de progresso
    document.getElementById('uploadProgressBar').classList.remove('hidden');

    showStatus(uploadStatus, 'Enviando e processando planilha, aguarde...', 'loading');

    try {
        // Simulação de progresso (substituir pela implementação real com XMLHttpRequest para progresso real)
        simulateProgress('uploadProgress', () => {
            // Esta parte seria substituída pelo fetch real
            fetch('/upload', {
                method: 'POST',
                body: formData
            })
            .then(resp => {
                if (!resp.ok) {
                    throw new Error(`Erro no servidor: ${resp.status}`);
                }
                return resp.json();
            })
            .then(data => {
                fileId = data.fileId;

                unidadeSelect.innerHTML = `<option value="">Todas as unidades</option>` +
                    data.unidades.map(u => `<option value="${u}">${u}</option>`).join('');

                filtersSection.classList.add('visible');
                filtersPlaceholder.style.display = 'none';

                showStatus(uploadStatus, 'Planilha carregada com sucesso! Agora selecione uma unidade para filtrar.', 'success');
            })
            .catch(error => {
                console.error('Erro:', error);
                showStatus(uploadStatus, `Falha ao carregar planilha: ${error.message}`, 'error');
            })
            .finally(() => {
                uploadBtn.disabled = false;
                uploadBtn.innerHTML = '<i class="fas fa-upload"></i><span>Carregar Planilha</span>';
                document.getElementById('uploadProgressBar').classList.add('hidden');
            });
        });

    } catch (error) {
        console.error('Erro:', error);
        showStatus(uploadStatus, `Falha ao carregar planilha: ${error.message}`, 'error');
        uploadBtn.disabled = false;
        uploadBtn.innerHTML = '<i class="fas fa-upload"></i><span>Carregar Planilha</span>';
        document.getElementById('uploadProgressBar').classList.add('hidden');
    }
});

// Função para simular progresso (substituir por implementação real)
function simulateProgress(progressElementId, callback) {
    const progressElement = document.getElementById(progressElementId);
    let width = 0;
    const interval = setInterval(() => {
        if (width >= 100) {
            clearInterval(interval);
            callback();
        } else {
            width += 5;
            progressElement.style.width = width + '%';
        }
    }, 100);
}

// Envio do formulário de upload múltiplo
formUploadMultiple.addEventListener('submit', async e => {
    e.preventDefault();

    if (!arquivosInput.files.length) {
        showStatus(uploadMultipleStatus, 'Por favor, selecione pelo menos um arquivo.', 'error');
        return;
    }

    const formData = new FormData();
    for (let i = 0; i < arquivosInput.files.length; i++) {
        formData.append('arquivos', arquivosInput.files[i]);
    }

    uploadMultipleBtn.disabled = true;
    uploadMultipleBtn.innerHTML = '<div class="spinner"></div><span>Processando...</span>';

    // Mostrar barra de progresso
    document.getElementById('uploadMultipleProgressBar').classList.remove('hidden');

    showStatus(uploadMultipleStatus, 'Enviando e processando planilhas, aguarde...', 'loading');

    try {
        // Simulação de progresso
        simulateProgress('uploadMultipleProgress', () => {
            // Esta parte seria substituída pelo fetch real
            fetch('/upload-multiple', {
                method: 'POST',
                body: formData
            })
            .then(resp => {
                if (!resp.ok) {
                    throw new Error(`Erro no servidor: ${resp.status}`);
                }
                return resp.json();
            })
            .then(data => {
                arquivosMultiplos = data;

                // Exibir lista de arquivos carregados
                arquivosCarregados.innerHTML = '';
                data.forEach((arquivo, index) => {
                    const div = document.createElement('div');
                    div.className = 'arquivo-item';
                    div.innerHTML = `
                        <div class="arquivo-info">
                            <i class="fas fa-file-excel" style="color: #1D6F42;"></i>
                            <span class="arquivo-nome">${arquivo.fileName}</span>
                        </div>
                        <div class="arquivo-unidades">
                            <select id="unidade-${index}" class="unidade-select">
                                <option value="">Todas as unidades</option>
                                ${arquivo.unidades.map(u => `<option value="${u}">${u}</option>`).join('')}
                            </select>
                        </div>
                    `;
                    arquivosCarregados.appendChild(div);
                });

                arquivosLista.classList.remove('hidden');
                mergeSection.classList.add('visible');
                mergePlaceholder.style.display = 'none';

                showStatus(uploadMultipleStatus, 'Planilhas carregadas com sucesso! Selecione as unidades para cada arquivo.', 'success');
            })
            .catch(error => {
                console.error('Erro:', error);
                showStatus(uploadMultipleStatus, `Falha ao carregar planilhas: ${error.message}`, 'error');
            })
            .finally(() => {
                uploadMultipleBtn.disabled = false;
                uploadMultipleBtn.innerHTML = '<i class="fas fa-upload"></i><span>Carregar Planilhas</span>';
                document.getElementById('uploadMultipleProgressBar').classList.add('hidden');
            });
        });

    } catch (error) {
        console.error('Erro:', error);
        showStatus(uploadMultipleStatus, `Falha ao carregar planilhas: ${error.message}`, 'error');
        uploadMultipleBtn.disabled = false;
        uploadMultipleBtn.innerHTML = '<i class="fas fa-upload"></i><span>Carregar Planilhas</span>';
        document.getElementById('uploadMultipleProgressBar').classList.add('hidden');
    }
});

// Pré-visualização dos dados (single)
previewBtn.addEventListener('click', async () => {
    const unidade = unidadeSelect.value;

    if (!unidade) {
        showStatus(downloadStatus, 'Por favor, selecione uma unidade primeiro.', 'error');
        return;
    }

    previewBtn.disabled = true;
    previewBtn.innerHTML = '<div class="spinner"></div><span>Carregando...</span>';

    try {
        const resp = await fetch('/preview', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ fileId, unidade })
        });

        if (!resp.ok) {
            throw new Error(`Erro no servidor: ${resp.status}`);
        }

        const data = await resp.json();

        previewTable.innerHTML = '';

        if (data.cabecalho && data.cabecalho.length) {
            const headerRow = document.createElement('tr');
            data.cabecalho.forEach(h => {
                const th = document.createElement('th');
                th.textContent = h;
                headerRow.appendChild(th);
            });
            previewTable.appendChild(headerRow);

            if (data.preview && data.preview.length) {
                data.preview.forEach(row => {
                    const tr = document.createElement('tr');
                    data.cabecalho.forEach(h => {
                        const td = document.createElement('td');
                        td.textContent = row[h] || '';
                        tr.appendChild(td);
                    });
                    previewTable.appendChild(tr);
                });

                previewCount.textContent = `${data.preview.length} linhas`;
            } else {
                const tr = document.createElement('tr');
                const td = document.createElement('td');
                td.colSpan = data.cabecalho.length;
                td.textContent = 'Nenhum dado encontrado.';
                td.style.textAlign = 'center';
                tr.appendChild(td);
                previewTable.appendChild(tr);

                previewCount.textContent = '0 linhas';
            }
        } else {
            previewTable.innerHTML = '<tr><td style="text-align: center; padding: 20px;">Nenhum dado disponível para visualização.</td></tr>';
            previewCount.textContent = '0 linhas';
        }

        previewTable.scrollIntoView({ behavior: 'smooth', block: 'start' });

    } catch (error) {
        console.error('Erro:', error);
        showStatus(downloadStatus, `Falha ao carregar pré-visualização: ${error.message}`, 'error');
    } finally {
        previewBtn.disabled = false;
        previewBtn.innerHTML = '<i class="fas fa-eye"></i><span>Pré-visualizar</span>';
    }
});

// Download da planilha filtrada (single)
baixarBtn.addEventListener('click', async () => {
    const unidade = unidadeSelect.value;

    if (!unidade) {
        showStatus(downloadStatus, 'Por favor, selecione uma unidade primeiro.', 'error');
        return;
    }

    baixarBtn.disabled = true;
    baixarBtn.innerHTML = '<div class="spinner"></div><span>Gerando planilha...</span>';
    showStatus(downloadStatus, 'Filtrando e gerando planilha, aguarde...', 'loading');

    try {
        const resp = await fetch('/filtrar', {
            method:'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ fileId, unidade })
        });

        if (!resp.ok) {
            throw new Error(`Erro no servidor: ${resp.status}`);
        }

        const blob = await resp.blob();
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = `planilha_filtrada_${unidade || 'todas'}.xlsx`;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);

        showStatus(downloadStatus, 'Planilha baixada com sucesso!', 'success');
    } catch (error) {
        console.error('Erro:', error);
        showStatus(downloadStatus, `Falha ao baixar planilha: ${error.message}`, 'error');
    } finally {
        baixarBtn.disabled = false;
        baixarBtn.innerHTML = '<i class="fas fa-download"></i><span>Baixar Planilha Filtrada</span>';
    }
});

// Mesclar planilhas (multiple)
mesclarBtn.addEventListener('click', async () => {
    if (arquivosMultiplos.length === 0) {
        showStatus(mergeStatus, 'Nenhum arquivo carregado para mesclar.', 'error');
        return;
    }

    const arquivosParaMesclar = [];

    for (let i = 0; i < arquivosMultiplos.length; i++) {
        const select = document.getElementById(`unidade-${i}`);
        const unidade = select.value;

        if (!unidade) {
            showStatus(mergeStatus, 'Por favor, selecione uma unidade para cada arquivo.', 'error');
            return;
        }

        arquivosParaMesclar.push({
            fileId: arquivosMultiplos[i].fileId,
            unidade: unidade
        });
    }

    mesclarBtn.disabled = true;
    mesclarBtn.innerHTML = '<div class="spinner"></div><span>Mesclando...</span>';
    showStatus(mergeStatus, 'Mesclando planilhas, aguarde...', 'loading');

    try {
        const resp = await fetch('/mesclar', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ arquivos: arquivosParaMesclar })
        });

        if (!resp.ok) {
            throw new Error(`Erro no servidor: ${resp.status}`);
        }

        const blob = await resp.blob();
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = `planilhas_mescladas.xlsx`;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);

        showStatus(mergeStatus, 'Planilhas mescladas e baixadas com sucesso!', 'success');
    } catch (error) {
        console.error('Erro:', error);
        showStatus(mergeStatus, `Falha ao mesclar planilhas: ${error.message}`, 'error');
    } finally {
        mesclarBtn.disabled = false;
        mesclarBtn.innerHTML = '<i class="fas fa-object-group"></i><span>Mesclar Planilhas</span>';
    }
});

// Função auxiliar para mostrar mensagens de status
function showStatus(element, message, type) {
    element.textContent = message;
    element.classList.remove('hidden', 'success', 'error', 'loading');
    element.classList.add(type, 'visible');

    if (type === 'success') {
        setTimeout(() => {
            element.classList.add('hidden');
            element.classList.remove('visible');
        }, 5000);
    }
}
// Bahia é o mundo