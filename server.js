const express = require('express');
const multer = require('multer');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

const app = express();
const upload = multer({ dest: 'uploads/' });

app.use(express.static('public'));
app.use(express.urlencoded({ extended: true }));
app.use(express.json());
let teste
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
        res.json({ unidades, filePath: req.file.path, fileName: req.file.originalname });
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
            resultados.push({
                fileName: file.originalname,
                filePath: file.path,
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
    const { filePath, unidade } = req.body;
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
    const { filePath, unidade } = req.body;
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
            
            // Label
            ws.getCell(`A${rowNumber}`).value = label + ':';
            ws.getCell(`A${rowNumber}`).font = { bold: true };
            ws.getCell(`A${rowNumber}`).fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'F2F2F2' }
            };
            ws.getCell(`A${rowNumber}`).border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
            
            // Valor
            ws.getCell(`B${rowNumber}`).value = infoValues[i];
            ws.getCell(`B${rowNumber}`).border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
            
            // Mesclar células se necessário para valores longos
            if (label === 'Objeto' && infoValues[i].length > 50) {
                ws.mergeCells(`B${rowNumber}:F${rowNumber}`);
            } else {
                ws.mergeCells(`B${rowNumber}:C${rowNumber}`);
            }
        });

        // Espaço entre informações e tabela
        const dataStartRow = infoLabels.length + 5;

        // --- Cabeçalho da tabela ---
        const cabecalho = Object.keys(dados.itens[0]);
        const headerRow = ws.getRow(dataStartRow);
        
        cabecalho.forEach((h, i) => {
            const cell = headerRow.getCell(i + 1);
            cell.value = h;
            cell.font = { bold: true, color: { argb: 'FFFFFF' } };
            cell.alignment = { vertical: 'middle', horizontal: 'center' };
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: '1F4E78' }
            };
            cell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
        });

        // --- Dados da tabela ---
        dados.itens.forEach((item, rowIndex) => {
            const row = ws.getRow(dataStartRow + rowIndex + 1);
            
            cabecalho.forEach((h, colIndex) => {
                const cell = row.getCell(colIndex + 1);
                let value = item[h] || '';
                
                // Formatar números
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
                
                // Formatação visual
                cell.border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };
                
                // Zebra stripes
                if (rowIndex % 2 === 0) {
                    cell.fill = {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: 'F9F9F9' }
                    };
                }
            });
        });

        // --- Ajustar largura das colunas ---
        ws.columns.forEach((column, i) => {
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

        res.download(outputPath);

    } catch (err) {
        console.error(err);
        res.status(500).send('Erro ao gerar planilha.');
    }
});

// Rota para mesclar múltiplas planilhas
// Rota para mesclar múltiplas planilhas
app.post('/mesclar', async (req, res) => {
    const { arquivos } = req.body;
    
    try {
        const newWB = new ExcelJS.Workbook();
        const ws = newWB.addWorksheet('Dados Mesclados');

        let currentRow = 1;

        for (const arq of arquivos) {
            const { filePath, unidade } = arq;
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.readFile(filePath);
            const sheet = workbook.worksheets[0];

            const dados = filtrarDados(sheet, unidade);

            if (dados.itens.length === 0) continue;

            // Adicionar cabeçalho do bloco
            ws.mergeCells(`A${currentRow}:F${currentRow}`);
            ws.getCell(`A${currentRow}`).value = `Planilha: ${path.basename(filePath)} | Unidade: ${unidade}`;
            ws.getCell(`A${currentRow}`).font = { bold: true, size: 14 };
            ws.getCell(`A${currentRow}`).fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'E6E6E6' }
            };
            currentRow++;

            // Adicionar metadados COMPLETOS (incluindo número da ata)
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
                
                // Adicionar bordas para melhor visualização
                ws.getCell(`A${currentRow + i}`).border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };
                
                // Mesclar células para valores longos (especialmente objeto)
                if (i === 1 && md.length > 50) { // índice 1 é o objeto
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
                cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: '1F4E78' }
                };
                cell.border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };
            });
            currentRow++;

            // Adicionar linhas de dados
            dados.itens.forEach((item, rowIndex) => {
                const row = ws.getRow(currentRow);
                cabecalho.forEach((h, colIndex) => {
                    const cell = row.getCell(colIndex + 1);
                    let value = item[h] || '';

                    // Formatar números
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
                    
                    // Formatação visual
                    cell.border = {
                        top: { style: 'thin' },
                        left: { style: 'thin' },
                        bottom: { style: 'thin' },
                        right: { style: 'thin' }
                    };
                    
                    // Zebra stripes
                    if (rowIndex % 2 === 0) {
                        cell.fill = {
                            type: 'pattern',
                            pattern: 'solid',
                            fgColor: { argb: 'F9F9F9' }
                        };
                    }
                });
                currentRow++;
            });

            // Adicionar espaçamento entre blocos
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

        res.download(outputPath);
    } catch (err) {
        console.error(err);
        res.status(500).send('Erro ao mesclar planilhas.');
    }
});

// Função para filtrar dados da planilha
function filtrarDados(sheet, unidadeEscolhida) {
    let capturando = false;
    let cabecalho = [];
    const itens = [];
    const info = { 
        numeroAta: '', 
        objeto: '', 
        negociacao: '', 
        inicioVigencia: '', 
        finalVigencia: '' 
    };

    let encontrouCabecalho = false;

    // Extrair informações do cabeçalho de forma mais precisa e direta
    try {
        // Número da Ata - Linha 13, Coluna D
        info.numeroAta = sheet.getCell('D13').value || '';
        
        // Objeto - Linha 14, Coluna D
        info.objeto = sheet.getCell('D14').value || '';
        
        // Negociação - Linha 14, Coluna T
        info.negociacao = sheet.getCell('T14').value || '';
        
        // Início da Vigência - Linha 15, Coluna T
        info.inicioVigencia = formatarData(sheet.getCell('T15').value);
        
        // Final da Vigência - Linha 16, Coluna T
        info.finalVigencia = formatarData(sheet.getCell('T16').value);
        
    } catch (e) {
        console.error('Erro ao extrair informações do cabeçalho:', e);
    }

    // Função auxiliar para formatar datas
    function formatarData(valor) {
        if (!valor) return '';
        
        // Se já é uma string formatada, retorna como está
        if (typeof valor === 'string') {
            // Verifica se é uma data no formato brasileiro
            const dataRegex = /(\d{2}\/\d{2}\/\d{4})/;
            const match = valor.match(dataRegex);
            if (match) return match[1];
            return valor;
        }
        
        // Se é um objeto Date, formata para o padrão brasileiro
        if (valor instanceof Date) {
            return valor.toLocaleDateString('pt-BR');
        }
        
        // Se é um número (serial do Excel), converte para data
        if (typeof valor === 'number') {
            const data = new Date((valor - 25569) * 86400 * 1000);
            return data.toLocaleDateString('pt-BR');
        }
        
        return String(valor);
    }

    sheet.eachRow((row, rowNumber) => {
        const valores = row.values.slice(1);

        // Verificar se é uma linha de unidade
        const linhaUnidade = valores.find(v => typeof v === 'string' && v.trim().startsWith('SESC -'));
        if (linhaUnidade) {
            const unidadeLinha = linhaUnidade.trim();
            capturando = unidadeEscolhida === 'Todas' || unidadeLinha === unidadeEscolhida;
            encontrouCabecalho = false; // Resetar ao encontrar nova unidade
            return;
        }

        // Se estamos capturando dados para a unidade selecionada
        if (capturando) {
            // Procurar pelo cabeçalho da tabela
            if (!encontrouCabecalho) {
                const temCabecalho = valores.some(v => 
                    v && typeof v === 'string' && 
                    (v.toLowerCase().includes('descrição') || 
                     v.toLowerCase().includes('item') ||
                     v.toLowerCase().includes('código'))
                );
                
                if (temCabecalho) {
                    cabecalho = valores.map(v => v || '');
                    encontrouCabecalho = true;
                    return;
                }
            } else {
                // Se já encontramos o cabeçalho, capturar os dados
                const textoLinha = valores.map(v => v ? v.toString() : '').join(' ').trim();
                const linhaVazia = valores.every(v => v === null || v === '' || v.toString().trim() === '');
                const ehRodape = textoLinha.toLowerCase().includes('itens por unidade') || 
                                textoLinha.toLowerCase().includes('total') || 
                                textoLinha.toLowerCase().includes('observações') ||
                                textoLinha.toLowerCase().includes('subtotal');
                
                if (!linhaVazia && !ehRodape) {
                    const item = {};
                    cabecalho.forEach((coluna, index) => {
                        if (coluna) {
                            item[coluna] = valores[index] || '';
                        }
                    });
                    
                    // Só adicionar se pelo menos uma célula tem valor
                    if (Object.values(item).some(v => v !== '')) {
                        itens.push(item);
                    }
                }
            }
        }
    });

    return { info, itens };
}

// Rota para limpar arquivos temporários
app.post('/limpar', (req, res) => {
    const diretorio = 'uploads/';
    
    fs.readdir(diretorio, (err, files) => {
        if (err) {
            console.error(err);
            return res.status(500).send('Erro ao limpar arquivos.');
        }
        
        for (const file of files) {
            if (file !== '.gitkeep') { // Não excluir arquivo .gitkeep se existir
                fs.unlink(path.join(diretorio, file), err => {
                    if (err) console.error(err);
                });
            }
        }
        
        res.send('Arquivos temporários removidos.');
    });
});

// Iniciar servidor
const PORT = process.env.PORT || 3000; // Render vai injetar process.env.PORT
app.listen(PORT, () => console.log(`Servidor rodando na porta ${PORT}`));