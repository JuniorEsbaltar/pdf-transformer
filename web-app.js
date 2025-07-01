document.addEventListener('DOMContentLoaded', () => {
    // Elementos de interface
    const sourceUploadArea = document.getElementById('sourceUploadArea');
    const sourceFileInput = document.getElementById('sourceFileInput');
    const sourceFileList = document.getElementById('sourceFileList');
    const controleUploadArea = document.getElementById('controleUploadArea');
    const controleFileInput = document.getElementById('controleFileInput');
    const controleFileList = document.getElementById('controleFileList');
    const processButton = document.getElementById('processButton');
    const logOutput = document.getElementById('logOutput');
    const downloadSection = document.getElementById('downloadSection');
    const downloadLink = document.getElementById('downloadLink');
    const downloadReportLink = document.getElementById('downloadReportLink');

    // Armazenamento de arquivos
    const filesData = {
        source: [],
        controle: null
    };

    // Configurar eventos de upload
    sourceUploadArea.addEventListener('click', () => sourceFileInput.click());
    controleUploadArea.addEventListener('click', () => controleFileInput.click());

    // Configurar drag & drop para arquivos source
    setupDragDrop(sourceUploadArea, 'source');
    setupDragDrop(controleUploadArea, 'controle');

    // Processar arquivos selecionados via input
    sourceFileInput.addEventListener('change', event => handleFiles(event.target.files, 'source'));
    controleFileInput.addEventListener('change', event => handleFiles(event.target.files, 'controle'));

    // Configurar botão de processamento
    processButton.addEventListener('click', processFiles);

    // Função para configurar drag & drop
    function setupDragDrop(element, fileType) {
        element.addEventListener('dragover', e => {
            e.preventDefault();
            e.stopPropagation();
            element.classList.add('bg-light');
        });

        element.addEventListener('dragleave', e => {
            e.preventDefault();
            e.stopPropagation();
            element.classList.remove('bg-light');
        });

        element.addEventListener('drop', e => {
            e.preventDefault();
            e.stopPropagation();
            element.classList.remove('bg-light');
            
            const files = e.dataTransfer.files;
            handleFiles(files, fileType);
        });
    }

    // Função para processar arquivos selecionados
    function handleFiles(files, fileType) {
        if (fileType === 'source') {
            // Para arquivos source, adicionar à lista
            Array.from(files).forEach(file => {
                if (file.name.endsWith('.xlsx')) {
                    // Verificar se o arquivo já existe
                    const exists = filesData.source.find(f => f.name === file.name);
                    if (!exists) {
                        filesData.source.push(file);
                    }
                }
            });
            updateFileList();
        } else if (fileType === 'controle') {
            // Para o arquivo de controle, substituir o atual
            if (files.length > 0 && files[0].name.endsWith('.xlsx')) {
                filesData.controle = files[0];
                updateFileList();
            }
        }
        
        // Verificar se podemos habilitar o botão de processamento
        checkProcessButtonState();
    }

    // Atualizar a visualização da lista de arquivos
    function updateFileList() {
        // Limpar listas
        sourceFileList.innerHTML = '';
        controleFileList.innerHTML = '';
        
        // Adicionar arquivos source
        filesData.source.forEach(file => {
            const item = document.createElement('div');
            item.className = 'file-item';
            
            const name = document.createElement('span');
            name.textContent = file.name;
            
            const removeBtn = document.createElement('button');
            removeBtn.className = 'btn btn-sm btn-danger';
            removeBtn.textContent = 'Remover';
            removeBtn.onclick = () => {
                filesData.source = filesData.source.filter(f => f !== file);
                updateFileList();
                checkProcessButtonState();
            };
            
            item.appendChild(name);
            item.appendChild(removeBtn);
            sourceFileList.appendChild(item);
        });
        
        // Adicionar arquivo de controle
        if (filesData.controle) {
            const item = document.createElement('div');
            item.className = 'file-item';
            
            const name = document.createElement('span');
            name.textContent = filesData.controle.name;
            
            const removeBtn = document.createElement('button');
            removeBtn.className = 'btn btn-sm btn-danger';
            removeBtn.textContent = 'Remover';
            removeBtn.onclick = () => {
                filesData.controle = null;
                updateFileList();
                checkProcessButtonState();
            };
            
            item.appendChild(name);
            item.appendChild(removeBtn);
            controleFileList.appendChild(item);
        }
    }
    
    // Verificar se podemos habilitar o botão de processamento
    function checkProcessButtonState() {
        processButton.disabled = filesData.source.length === 0 || !filesData.controle;
    }

    // Função para extrair o ID de referência do aplicativo
    function extrairIdReferencia(referenciaCompleta) {
        // Verificar se a referência existe e é uma string
        if (!referenciaCompleta || typeof referenciaCompleta !== 'string') {
            return '';
        }
        
        // Procurar pelo hífen e extrair o que vem depois
        const partes = referenciaCompleta.split('-');
        
        // Se houver pelo menos um hífen, retornar a última parte
        if (partes.length > 1) {
            return partes[partes.length - 1].trim();
        }
        
        // Se não houver hífen, retornar a string original
        return referenciaCompleta.trim();
    }

    // Função para adicionar log à interface
    function log(message, type = 'info') {
        const logLine = document.createElement('div');
        logLine.className = `log-line ${type}`;
        logLine.textContent = message;
        logOutput.appendChild(logLine);
        logOutput.scrollTop = logOutput.scrollHeight;
    }

    // Função para processar os arquivos Excel e extrair as informações necessárias
    async function processarArquivosExcel(arquivos) {
        log(`Processando ${arquivos.length} arquivos Excel...`);
        
        // Array para armazenar todos os dados extraídos
        const dadosExtraidos = [];
        
        // Processar cada arquivo
        for (const arquivo of arquivos) {
            log(`Processando arquivo: ${arquivo.name}`);
            
            try {
                // Ler o arquivo Excel
                const arrayBuffer = await arquivo.arrayBuffer();
                const workbook = XLSX.read(arrayBuffer, { type: 'array' });
                const sheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[sheetName];
                const dados = XLSX.utils.sheet_to_json(worksheet);
                
                log(`Extraindo dados de ${dados.length} linhas do arquivo ${arquivo.name}`);
                
                // Extrair informações de cada linha
                const dadosArquivo = dados.map(linha => {
                    // Identificar o tipo de arquivo (omnie_1 ou omnie_2)
                    const isOmnie1 = arquivo.name.includes('omnie_1');
                    
                    // Extrair o ID de referência
                    const idCompleto = linha['Referência do Aplicativo'] || '';
                    const idReferencia = extrairIdReferencia(idCompleto);
                    
                    // Extrair o mês de referência e valor da comissão com base no tipo de arquivo
                    let mesReferencia, valorComissao;
                    
                    // Função para converter número serial do Excel para objeto Date
                    function excelDateToJSDate(excelDate) {
                        if (!excelDate || isNaN(excelDate)) return null;
                        const EXCEL_EPOCH = new Date(1899, 11, 30); // 30/12/1899
                        const millisecondsPerDay = 24 * 60 * 60 * 1000;
                        return new Date(EXCEL_EPOCH.getTime() + excelDate * millisecondsPerDay);
                    }
                    
                    if (isOmnie1) {
                        // Para omnie_1.xlsx - mês do recebimento e valor do repasse
                        mesReferencia = excelDateToJSDate(linha['Recebto.']);
                        valorComissao = linha['Valor do Repasse'] || 0;
                    } else {
                        // Para omnie_2.xlsx - mês da emissão e valor da comissão
                        mesReferencia = excelDateToJSDate(linha['Emissão']);
                        valorComissao = linha['Comissão R$'] || 0;
                    }
                    
                    // Formatar o mês como string (ex: "Jan", "Fev", etc.)
                    const meses = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'];
                    const mesFormatado = mesReferencia ? meses[mesReferencia.getMonth()] : 'N/A';
                    
                    return {
                        idReferencia,
                        idCompleto,
                        mesReferencia,
                        mesFormatado,
                        valorComissao: parseFloat(valorComissao),
                        fonte: arquivo.name,
                        linhaOriginal: linha
                    };
                });
                
                // Adicionar os dados extraídos ao array principal
                dadosExtraidos.push(...dadosArquivo);
                
            } catch (erro) {
                log(`Erro ao processar o arquivo ${arquivo.name}: ${erro.message}`, 'error');
                console.error(erro);
            }
        }
        
        log(`Total de registros processados: ${dadosExtraidos.length}`, 'success');
        return dadosExtraidos;
    }

    // Função para atualizar a planilha de controle com os dados extraídos
    async function atualizarPlanilhaControle(registros, arquivoControle) {
        log('Atualizando planilha de controle...');
        
        try {
            // Ler o arquivo de controle
            const arrayBuffer = await arquivoControle.arrayBuffer();
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.load(arrayBuffer);
            
            const sheetName = "APPs";
            const worksheet = workbook.getWorksheet(sheetName);
            if (!worksheet) {
                throw new Error(`Planilha "${sheetName}" não encontrada na planilha de controle.`);
            }
            
            // Obter os cabeçalhos e criar mapa de colunas
            const headerRow = worksheet.getRow(1);
            const headers = [];
            const colunasPorNome = {};
            
            headerRow.eachCell({ includeEmpty: false }, (cell, colNumber) => {
                headers[colNumber] = cell.value;
                if (cell.value) {
                    colunasPorNome[cell.value] = colNumber;
                }
            });
            
            // Identificar a coluna de IDs
            let colunaId = null;
            const possiveisColunasId = ['REFERÊNCIA APLICATIVO', 'ID', 'Referência', 'Referência do Aplicativo'];

            for (const colunaPossivel of possiveisColunasId) {
                if (colunasPorNome[colunaPossivel]) {
                    colunaId = colunaPossivel;
                    log(`Coluna de IDs encontrada: ${colunaId}`);
                    break;
                }
            }

            if (!colunaId) {
                throw new Error('Não foi possível encontrar a coluna de IDs de referência');
            }

            // Mapear os meses para os nomes das colunas na planilha de controle
            const mapeamentoMeses = {
                'Jan': 'Recorrência Janeiro',
                'Fev': 'Recorrência Fevereiro',
                'Mar': 'Recorrência Março',
                'Abr': 'Recorrência Abril',
                'Mai': 'Recorrência Maio',
                'Jun': 'Recorrência Junho',
                'Jul': 'Recorrência Julho',
                'Ago': 'Recorrência Agosto',
                'Set': 'Recorrência Setembro',
                'Out': 'Recorrência Outubro',
                'Nov': 'Recorrência Novembro',
                'Dez': 'Recorrência Dezembro'
            };
            
            // Armazenar as células que foram alteradas para o relatório
            const celulasAlteradas = [];
            
            // Contador de atualizações realizadas
            let atualizacoes = 0;
            let naoEncontrados = [];
            
            // Para cada registro válido, atualizar a planilha de controle
            for (const registro of registros) {
                const { idReferencia, mesFormatado, valorComissao } = registro;
                
                if (!idReferencia || !mesFormatado || !valorComissao) {
                    continue; // Pular registros inválidos
                }
                
                // Verificar se o mês está mapeado
                const colunaMes = mapeamentoMeses[mesFormatado];
                if (!colunaMes) {
                    log(`Mês não mapeado: ${mesFormatado}`, 'error');
                    continue;
                }
                
                // Verificar se a coluna existe na planilha
                let colunaIndex = colunasPorNome[colunaMes];
                if (!colunaIndex) {
                    const colunasRelacionadas = Object.keys(colunasPorNome).filter(coluna => 
                        coluna.toLowerCase().includes(mesFormatado.toLowerCase()) || 
                        (mesFormatado === 'Jun' && coluna.toLowerCase().includes('junho')) ||
                        (mesFormatado === 'Mai' && coluna.toLowerCase().includes('maio'))
                    );
                    
                    if (colunasRelacionadas.length > 0) {
                        log(`Coluna alternativa encontrada para ${mesFormatado}: ${colunasRelacionadas[0]}`);
                        colunaIndex = colunasPorNome[colunasRelacionadas[0]];
                    }
                }
                
                if (!colunaIndex) {
                    log(`Coluna "${colunaMes}" não encontrada na planilha de controle`, 'warning');
                    continue;
                }
                
                // Buscar a linha com o ID de referência
                let linhaEncontrada = -1;
                const colunaIdIndex = colunasPorNome[colunaId];
                
                if (colunaIdIndex) {
                    log(`Buscando ID ${idReferencia} na coluna ${colunaId}...`);
                    
                    // Buscar em todas as linhas da planilha
                    let linhaAtual = 2; // Começar após o cabeçalho
                    let celulaEncontrada = false;
                    
                    while (!celulaEncontrada && linhaAtual <= worksheet.rowCount) {
                        const cell = worksheet.getRow(linhaAtual).getCell(colunaIdIndex);
                        let valorCelula = '';
                        
                        if (cell.text) valorCelula = cell.text;
                        else if (cell.value) {
                            valorCelula = cell.value.toString();
                        }
                        
                        // Verificar se o valor da célula contém o ID procurado
                        if (valorCelula && valorCelula.includes(idReferencia)) {
                            log(`ID ${idReferencia} encontrado na linha ${linhaAtual}`);
                            linhaEncontrada = linhaAtual;
                            celulaEncontrada = true;
                            break;
                        }
                        
                        linhaAtual++;
                    }
                }
                
                if (linhaEncontrada === -1) {
                    naoEncontrados.push(idReferencia);
                    log(`ID de referência "${idReferencia}" não encontrado na planilha de controle`, 'warning');
                    continue;
                }
                
                // Obter o valor atual da célula
                const cell = worksheet.getRow(linhaEncontrada).getCell(colunaIndex);
                let valorAtual = cell.value;
                
                // Converter para número para comparação
                let valorAtualNumerico = 0;
                if (valorAtual !== null && valorAtual !== undefined) {
                    if (typeof valorAtual === 'number') {
                        valorAtualNumerico = valorAtual;
                    } else if (typeof valorAtual === 'string' && !isNaN(valorAtual)) {
                        valorAtualNumerico = parseFloat(valorAtual);
                    } else if (typeof valorAtual === 'object' && valorAtual.result) {
                        valorAtualNumerico = parseFloat(valorAtual.result);
                    }
                }
                
                // Se o valor atual for diferente do novo valor, atualizar
                if (Math.abs(valorAtualNumerico - valorComissao) > 0.01) {
                    log(`Atualizando ${idReferencia} - ${mesFormatado}: ${valorAtual || 'vazio'} -> ${valorComissao}`);
                    
                    // Atualizar o valor na célula
                    cell.value = valorComissao;
                    
                    // Registrar a célula alterada para o relatório
                    celulasAlteradas.push({
                        idReferencia,
                        mesReferencia: mesFormatado,
                        colunaMes,
                        valorAntigo: valorAtual || 'vazio',
                        valorNovo: valorComissao,
                        celula: `${worksheet.getColumn(colunaIndex).letter}${linhaEncontrada}`
                    });
                    
                    atualizacoes++;
                } else {
                    log(`Valor já atualizado para ${idReferencia} - ${mesFormatado}: ${valorComissao}`);
                }
            }
            
            // Exibir IDs não encontrados
            if (naoEncontrados.length > 0) {
                log(`Os seguintes IDs não foram encontrados na planilha de controle:`, 'warning');
                log(naoEncontrados.join(', '), 'warning');
            }
            
            // Criar um relatório de alterações em formato texto
            let relatorioConteudo = `RELATÓRIO DE ALTERAÇÕES - ${new Date().toLocaleString()}\n\n`;
            relatorioConteudo += `Total de registros processados: ${registros.length}\n`;
            relatorioConteudo += `Total de atualizações realizadas: ${atualizacoes}\n\n`;
            
            if (celulasAlteradas.length > 0) {
                relatorioConteudo += `Detalhamento das alterações:\n`;
                celulasAlteradas.forEach((celula, index) => {
                    relatorioConteudo += `${index + 1}. ID: ${celula.idReferencia} - Mês: ${celula.mesReferencia} (Célula: ${celula.celula})\n`;
                    relatorioConteudo += `   Valor anterior: ${celula.valorAntigo}\n`;
                    relatorioConteudo += `   Valor atual: ${celula.valorNovo}\n\n`;
                });
            }
            
            if (naoEncontrados.length > 0) {
                relatorioConteudo += `\nIDs não encontrados na planilha de controle:\n`;
                naoEncontrados.forEach((id, index) => {
                    relatorioConteudo += `${index + 1}. ${id}\n`;
                });
            }
            
            // Salvar a planilha em formato de dados binários
            const buffer = await workbook.xlsx.writeBuffer();
            
            if (atualizacoes > 0) {
                log(`Planilha atualizada com sucesso! ${atualizacoes} atualizações realizadas.`, 'success');
            } else {
                log('Nenhuma atualização necessária na planilha de controle.', 'info');
            }
            
            return {
                planilhaAtualizada: buffer,
                relatorioAlteracoes: relatorioConteudo
            };
            
        } catch (erro) {
            log(`Erro ao atualizar a planilha de controle: ${erro.message}`, 'error');
            console.error(erro);
            throw erro;
        }
    }

    // Função principal para processar os arquivos
    async function processFiles() {
        // Limpar logs anteriores
        logOutput.innerHTML = '';
        downloadSection.style.display = 'none';
        
        // Desabilitar botão durante o processamento
        processButton.disabled = true;
        processButton.innerHTML = `<span class="spinner-border spinner-border-sm"></span> Processando...`;
        
        try {
            log('Iniciando processamento de arquivos...', 'info');
            
            // Verificar se temos todos os arquivos necessários
            if (filesData.source.length === 0) {
                throw new Error('Nenhum arquivo source selecionado');
            }
            if (!filesData.controle) {
                throw new Error('Nenhuma planilha de controle selecionada');
            }
            
            // Processar arquivos source
            log('Processando arquivos source...', 'info');
            const dadosExtraidos = await processarArquivosExcel(filesData.source);
            
            // Filtrar registros válidos
            log('Filtrando registros válidos...', 'info');
            const registrosValidos = dadosExtraidos.filter(item => 
                item.idReferencia && 
                item.mesReferencia && 
                !isNaN(item.valorComissao) && 
                item.valorComissao > 0
            );
            
            log(`Total de registros válidos: ${registrosValidos.length}`, 'success');
            
            // Atualizar planilha de controle
            log('Atualizando planilha de controle...', 'info');
            const { planilhaAtualizada, relatorioAlteracoes } = await atualizarPlanilhaControle(registrosValidos, filesData.controle);
            
            // Criar link para download da planilha atualizada
            const planilhaBlob = new Blob([planilhaAtualizada], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
            const planilhaUrl = URL.createObjectURL(planilhaBlob);
            downloadLink.href = planilhaUrl;
            downloadLink.download = 'controle_planilha_atualizado.xlsx';
            
            // Criar link para download do relatório
            const relatorioBlob = new Blob([relatorioAlteracoes], { type: 'text/plain' });
            const relatorioUrl = URL.createObjectURL(relatorioBlob);
            downloadReportLink.href = relatorioUrl;
            downloadReportLink.download = 'relatorio_alteracoes.txt';
            
            // Exibir seção de download
            downloadSection.style.display = 'block';
            
            log('Processamento concluído com sucesso!', 'success');
            
        } catch (erro) {
            log(`Erro: ${erro.message}`, 'error');
            console.error(erro);
        } finally {
            // Restaurar botão
            processButton.disabled = false;
            processButton.textContent = 'Processar Planilhas';
        }
    }
})
