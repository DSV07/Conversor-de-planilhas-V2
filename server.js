const express = require('express');
const multer = require('multer');
const ExcelJS = require('exceljs');
const path =require('path');
const fs = require('fs');
const crypto = require('crypto');

const app = express();
const upload = multer({ dest: 'uploads/' });

// Armazenamento em memória para mapear IDs de arquivo para caminhos de arquivo
const fileStore = {};

app.use(express.static('public'));
app.use(express.urlencoded({ extended: true }));
app.use(express.json());

// Colunas que sempre devem ser números
const colunasNumericas = ['Valor', 'Saldo', 'Inicial', 'Solicitada', 'Consumida', 'Saldo Atual'];

// Rota para upload de arquivo único
app.post('/upload', upload.single('arquivo'), async (req, res) => {
    try {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(req.file.path);
        const sheet = workbook.worksheets[0];

        const unidadesSet = new Set();
        sheet.eachRow(row => {
            row.eachCell(cell => {
                if (cell.value && typeof cell.value === 'string' && cell.value.trim().startsWith('SESC -')) {
                    unidadesSet.add(cell.value.trim());
                }
            });
        });

        const unidades = Array.from(unidadesSet).sort();

        // Gerar um ID seguro e armazenar o caminho do arquivo
        const fileId = crypto.randomUUID();
        fileStore[fileId] = { path: req.file.path, name: req.file.originalname };

        res.json({ unidades, fileId, fileName: req.file.originalname });
    } catch (err) {
        console.error(err);
        res.status(500).send('Erro ao processar arquivo.');
    }
});

// Rota para upload múltiplo
app.post('/upload-multiple', upload.array('arquivos'), async (req, res) => {
    try {
        const resultados = [];
        for (const file of req.files) {
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.readFile(file.path);
            const sheet = workbook.worksheets[0];

            const unidadesSet = new Set();
            sheet.eachRow(row => {
                row.eachCell(cell => {
                    if (cell.value && typeof cell.value === 'string' && cell.value.trim().startsWith('SESC -')) {
                        unidadesSet.add(cell.value.trim());
                    }
                });
            });

            const unidades = Array.from(unidadesSet).sort();

            // Gerar um ID seguro para cada arquivo
            const fileId = crypto.randomUUID();
            fileStore[fileId] = { path: file.path, name: file.originalname };

            resultados.push({
                fileName: file.originalname,
                fileId,
                unidades
            });
        }
        res.json(resultados);
    } catch (err) {
        console.error(err);
        res.status(500).send('Erro ao processar arquivos.');
    }
});

// Rota para pré-visualização
app.post('/preview', async (req, res) => {
    const { fileId, unidade } = req.body;

    const fileData = fileStore[fileId];
    if (!fileData) {
        return res.status(404).send('Arquivo não encontrado ou sessão expirada.');
    }
    const filePath = fileData.path;

    try {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(filePath);
        const sheet = workbook.worksheets[0];

        const dados = filtrarDados(sheet, unidade);

        res.json({ cabecalho: dados.itens.length ? Object.keys(dados.itens[0]) : [], preview: dados.itens.slice(0, 5) });
    } catch (err) {
        console.error(err);
        res.status(500).send('Erro ao gerar pré-visualização.');
    }
});

// Rota para filtrar um único arquivo
app.post('/filtrar', async (req, res) => {
    const { fileId, unidade } = req.body;

    const fileData = fileStore[fileId];
    if (!fileData) {
        return res.status(404).send('Arquivo não encontrado ou sessão expirada.');
    }
    const filePath = fileData.path;

    try {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(filePath);
        const sheet = workbook.worksheets[0];

        const dados = filtrarDados(sheet, unidade);

        if (!dados.itens.length) return res.status(400).send('Não foi possível filtrar dados.');

        const newWB = new ExcelJS.Workbook();
        const ws = newWB.addWorksheet(unidade || 'Dados Filtrados');

        // --- Título principal ---
        ws.mergeCells('A1:F1');
        ws.getCell('A1').value = 'RELATÓRIO DE DADOS FILTRADOS';
        ws.getCell('A1').font = { bold: true, size: 16 };
        ws.getCell('A1').alignment = { horizontal: 'center' };
        ws.getCell('A1').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'D9D9D9' }
        };

        // --- Informações gerais ---
        const infoLabels = ['Número da Ata', 'Objeto', 'Negociação', 'Início Vigência', 'Final Vigência'];
        const infoValues = [
            dados.info.numeroAta || '-',
            dados.info.objeto || '-',
            dados.info.negociacao || '-',
            dados.info.inicioVigencia || '-',
            dados.info.finalVigencia || '-'
        ];

        // Adicionar informações com formatação
        infoLabels.forEach((label, i) => {
            const rowNumber = i + 3;
            ws.getCell(`A${rowNumber}`).value = label + ':';
            ws.getCell(`A${rowNumber}`).font = { bold: true };
            ws.getCell(`A${rowNumber}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'F2F2F2' }};
            ws.getCell(`A${rowNumber}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' }};

            ws.getCell(`B${rowNumber}`).value = infoValues[i];
            ws.getCell(`B${rowNumber}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' }};

            if (label === 'Objeto' && infoValues[i].length > 50) {
                ws.mergeCells(`B${rowNumber}:F${rowNumber}`);
            } else {
                ws.mergeCells(`B${rowNumber}:C${rowNumber}`);
            }
        });

        const dataStartRow = infoLabels.length + 5;

        // --- Cabeçalho da tabela ---
        const cabecalho = Object.keys(dados.itens[0]);
        const headerRow = ws.getRow(dataStartRow);

        cabecalho.forEach((h, i) => {
            const cell = headerRow.getCell(i + 1);
            cell.value = h;
            cell.font = { bold: true, color: { argb: 'FFFFFF' } };
            cell.alignment = { vertical: 'middle', horizontal: 'center' };
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '1F4E78' }};
            cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' }};
        });

        // --- Dados da tabela ---
        dados.itens.forEach((item, rowIndex) => {
            const row = ws.getRow(dataStartRow + rowIndex + 1);
            cabecalho.forEach((h, colIndex) => {
                const cell = row.getCell(colIndex + 1);
                let value = item[h] || '';

                if (colunasNumericas.includes(h)) {
                    if (value) {
                        let num = value.toString().replace(/\s/g, '').replace(/,/g, '.').replace(/[^0-9.-]/g, '');
                        cell.value = !isNaN(Number(num)) ? Number(num) : 0;
                    } else {
                        cell.value = 0;
                    }
                    cell.alignment = { horizontal: 'right', vertical: 'middle' };
                    cell.numFmt = '#,##0.00';
                } else {
                    cell.value = value;
                    cell.alignment = { horizontal: 'left', vertical: 'middle' };
                }

                cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' }};

                if (rowIndex % 2 === 0) {
                    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'F9F9F9' }};
                }
            });
        });

        // --- Ajustar largura das colunas ---
        ws.columns.forEach((column) => {
            let maxLength = 0;
            column.eachCell({ includeEmpty: true }, (cell) => {
                let columnLength = cell.value ? cell.value.toString().length : 10;
                if (columnLength > maxLength) {
                    maxLength = columnLength;
                }
            });
            column.width = Math.min(maxLength + 2, 50);
        });

        // --- Data de geração ---
        const lastRow = dataStartRow + dados.itens.length + 3;
        ws.mergeCells(`A${lastRow}:F${lastRow}`);
        ws.getCell(`A${lastRow}`).value = `Gerado em: ${new Date().toLocaleString('pt-BR')}`;
        ws.getCell(`A${lastRow}`).font = { italic: true, color: { argb: '666666' } };
        ws.getCell(`A${lastRow}`).alignment = { horizontal: 'right' };

        // --- Salvar e enviar ---
        const outputPath = path.join('uploads', `filtrado_${Date.now()}.xlsx`);
        await newWB.xlsx.writeFile(outputPath);

        res.download(outputPath, `planilha_filtrada_${unidade || 'todas'}.xlsx`, (err) => {
            if (err) {
                console.error('Erro ao enviar o arquivo:', err);
            }
            // Limpa o arquivo gerado
            fs.unlink(outputPath, (unlinkErr) => {
                if (unlinkErr) console.error('Erro ao limpar arquivo gerado:', unlinkErr);
            });
            // Limpa o arquivo original
            fs.unlink(filePath, (unlinkErr) => {
                if (unlinkErr) console.error('Erro ao limpar arquivo original:', unlinkErr);
            });
            // Remove a referência do armazenamento
            delete fileStore[fileId];
        });

    } catch (err) {
        console.error(err);
        res.status(500).send('Erro ao gerar planilha.');
    }
});

// Rota para mesclar múltiplas planilhas
app.post('/mesclar', async (req, res) => {
    const { arquivos } = req.body;

    try {
        const newWB = new ExcelJS.Workbook();
        const ws = newWB.addWorksheet('Dados Mesclados');

        let currentRow = 1;

        for (const arq of arquivos) {
            const { fileId, unidade } = arq;

            const fileData = fileStore[fileId];
            if (!fileData) {
                // Pula arquivos não encontrados para não quebrar a mesclagem
                console.warn(`Arquivo com ID ${fileId} não encontrado. Pulando.`);
                continue;
            }
            const filePath = fileData.path;

            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.readFile(filePath);
            const sheet = workbook.worksheets[0];

            const dados = filtrarDados(sheet, unidade);

            if (dados.itens.length === 0) continue;

            // Adicionar cabeçalho do bloco
            ws.mergeCells(`A${currentRow}:F${currentRow}`);
            ws.getCell(`A${currentRow}`).value = `Planilha: ${fileData.name} | Unidade: ${unidade}`;
            ws.getCell(`A${currentRow}`).font = { bold: true, size: 14 };
            ws.getCell(`A${currentRow}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'E6E6E6' }};
            currentRow++;

            // Adicionar metadados
            const metadados = [
                `Número da Ata: ${dados.info.numeroAta || '-'}`,
                `Objeto: ${dados.info.objeto || '-'}`,
                `Negociação: ${dados.info.negociacao || '-'}`,
                `Início Vigência: ${dados.info.inicioVigencia || '-'}`,
                `Final Vigência: ${dados.info.finalVigencia || '-'}`
            ];

            metadados.forEach((md, i) => {
                ws.getCell(`A${currentRow + i}`).value = md;
                ws.getCell(`A${currentRow + i}`).font = { italic: true };
                ws.getCell(`A${currentRow + i}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' }};
                if (i === 1 && md.length > 50) {
                    ws.mergeCells(`A${currentRow + i}:F${currentRow + i}`);
                } else {
                    ws.mergeCells(`A${currentRow + i}:C${currentRow + i}`);
                }
            });
            currentRow += metadados.length + 1;

            // Adicionar tabela de dados
            const cabecalho = Object.keys(dados.itens[0]);
            const headerRow = ws.getRow(currentRow);

            cabecalho.forEach((h, i) => {
                const cell = headerRow.getCell(i + 1);
                cell.value = h;
                cell.font = { bold: true, color: { argb: 'FFFFFF' } };
                cell.alignment = { vertical: 'middle', horizontal: 'center' };
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '1F4E78' }};
                cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' }};
            });
            currentRow++;

            // Adicionar linhas de dados
            dados.itens.forEach((item, rowIndex) => {
                const row = ws.getRow(currentRow);
                cabecalho.forEach((h, colIndex) => {
                    const cell = row.getCell(colIndex + 1);
                    let value = item[h] || '';

                    if (colunasNumericas.includes(h)) {
                        if (value) {
                            let num = value.toString().replace(/\s/g, '').replace(/,/g, '.').replace(/[^0-9.-]/g, '');
                            cell.value = !isNaN(Number(num)) ? Number(num) : 0;
                            cell.alignment = { horizontal: 'right', vertical: 'middle' };
                            cell.numFmt = '#,##0.00';
                        } else {
                            cell.value = 0;
                        }
                    } else {
                        cell.value = value;
                        cell.alignment = { horizontal: 'left', vertical: 'middle' };
                    }

                    cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' }};

                    if (rowIndex % 2 === 0) {
                        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'F9F9F9' }};
                    }
                });
                currentRow++;
            });

            currentRow += 2;
        }

        // Ajustar largura das colunas
        ws.columns.forEach(column => {
            let maxLength = 0;
            column.eachCell({ includeEmpty: true }, (cell) => {
                let columnLength = cell.value ? cell.value.toString().length : 10;
                if (columnLength > maxLength) {
                    maxLength = columnLength;
                }
            });
            column.width = Math.min(maxLength + 2, 50);
        });

        // Adicionar data de geração
        ws.mergeCells(`A${currentRow}:F${currentRow}`);
        ws.getCell(`A${currentRow}`).value = `Gerado em: ${new Date().toLocaleString('pt-BR')}`;
        ws.getCell(`A${currentRow}`).font = { italic: true, color: { argb: '666666' } };
        ws.getCell(`A${currentRow}`).alignment = { horizontal: 'right' };

        const outputPath = path.join('uploads', `mesclado_${Date.now()}.xlsx`);
        await newWB.xlsx.writeFile(outputPath);

        res.download(outputPath, 'planilhas_mescladas.xlsx', (err) => {
            if (err) {
                console.error('Erro ao enviar o arquivo mesclado:', err);
            }
            // Limpa o arquivo gerado
            fs.unlink(outputPath, (unlinkErr) => {
                if (unlinkErr) console.error('Erro ao limpar arquivo mesclado:', unlinkErr);
            });

            // Limpa todos os arquivos originais utilizados na mesclagem
            for (const arq of arquivos) {
                const fileData = fileStore[arq.fileId];
                if (fileData) {
                    fs.unlink(fileData.path, (unlinkErr) => {
                        if (unlinkErr) console.error(`Erro ao limpar arquivo original ${fileData.name}:`, unlinkErr);
                    });
                    // Remove a referência do armazenamento
                    delete fileStore[arq.fileId];
                }
            }
        });
    } catch (err) {
        console.error(err);
        res.status(500).send('Erro ao mesclar planilhas.');
    }
});

// Função auxiliar para formatar datas
function formatarData(valor) {
    if (!valor) return '';
    if (typeof valor === 'string') {
        const dataRegex = /(\d{2}\/\d{2}\/\d{4})/;
        const match = valor.match(dataRegex);
        if (match) return match[1];
        return valor;
    }
    if (valor instanceof Date) {
        return valor.toLocaleDateString('pt-BR');
    }
    if (typeof valor === 'number') {
        const data = new Date((valor - 25569) * 86400 * 1000);
        return data.toLocaleDateString('pt-BR');
    }
    return String(valor);
}

// Função para extrair metadados da planilha de forma robusta
function extrairMetadados(sheet) {
    const info = {
        numeroAta: '',
        objeto: '',
        negociacao: '',
        inicioVigencia: '',
        finalVigencia: ''
    };
    const labelMapping = {
        'número da ata': 'numeroAta',
        'objeto': 'objeto',
        'negociação': 'negociacao',
        'início da vigência': 'inicioVigencia',
        'final da vigência': 'finalVigencia'
    };
    const labelsToFind = Object.keys(labelMapping);
    let foundCount = 0;

    for (let rowNumber = 1; rowNumber <= Math.min(30, sheet.rowCount); rowNumber++) {
        const row = sheet.getRow(rowNumber);
        if (foundCount === labelsToFind.length) break;
        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
            if (cell.value && typeof cell.value === 'string') {
                const cellText = cell.value.trim().toLowerCase();
                for (const label of labelsToFind) {
                    if (cellText.startsWith(label)) {
                        const infoKey = labelMapping[label];
                        if (!info[infoKey]) {
                            for (let i = colNumber + 1; i <= row.cellCount; i++) {
                                const valueCell = row.getCell(i);
                                if (valueCell.value) {
                                    let result = valueCell.value;
                                    if (result && typeof result === 'object' && result.richText) {
                                        result = result.richText.map(rt => rt.text).join('');
                                    }
                                    info[infoKey] = result;
                                    foundCount++;
                                    break;
                                }
                            }
                        }
                    }
                }
            }
        });
    }

    info.inicioVigencia = formatarData(info.inicioVigencia);
    info.finalVigencia = formatarData(info.finalVigencia);
    return info;
}

// Função para filtrar dados da planilha
function filtrarDados(sheet, unidadeEscolhida) {
    let capturando = false;
    let cabecalho = [];
    const itens = [];
    let encontrouCabecalho = false;
    const info = extrairMetadados(sheet);

    sheet.eachRow((row) => {
        const valores = row.values.slice(1);
        const linhaUnidade = valores.find(v => typeof v === 'string' && v.trim().startsWith('SESC -'));
        if (linhaUnidade) {
            const unidadeLinha = linhaUnidade.trim();
            capturando = unidadeEscolhida === 'Todas' || unidadeLinha === unidadeEscolhida;
            encontrouCabecalho = false;
            return;
        }
        if (capturando) {
            if (!encontrouCabecalho) {
                const temCabecalho = valores.some(v => v && typeof v === 'string' && (v.toLowerCase().includes('descrição') || v.toLowerCase().includes('item') || v.toLowerCase().includes('código')));
                if (temCabecalho) {
                    cabecalho = valores.map(v => v || '');
                    encontrouCabecalho = true;
                    return;
                }
            } else {
                const textoLinha = valores.map(v => v ? v.toString() : '').join(' ').trim();
                const linhaVazia = valores.every(v => v === null || v === '' || v.toString().trim() === '');
                const ehRodape = textoLinha.toLowerCase().includes('itens por unidade') || textoLinha.toLowerCase().includes('total') || textoLinha.toLowerCase().includes('observações') || textoLinha.toLowerCase().includes('subtotal');
                if (!linhaVazia && !ehRodape) {
                    const item = {};
                    cabecalho.forEach((coluna, index) => {
                        if (coluna) {
                            item[coluna] = valores[index] || '';
                        }
                    });
                    if (Object.values(item).some(v => v !== '')) {
                        itens.push(item);
                    }
                }
            }
        }
    });
    return { info, itens };
}

// Iniciar servidor
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Servidor rodando na porta ${PORT}`));