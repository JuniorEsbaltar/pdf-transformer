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
    // const bpoSwitch = document.getElementById('bpoSwitch');

    // Inicializar tooltips do Bootstrap
    const tooltipTriggerList = [].slice.call(document.querySelectorAll('[data-bs-toggle="tooltip"]'));
    tooltipTriggerList.map(function (tooltipTriggerEl) {
        return new bootstrap.Tooltip(tooltipTriggerEl);
    });

    // Armazenamento de arquivos
    const filesData = {
        source: [],
        controle: null,
        isBPO: false
    };

    // Event listener para o switch de BPO
    // bpoSwitch.addEventListener('change', () => {
    //     filesData.isBPO = bpoSwitch.checked;
    //     log(`Opção de BPO ${bpoSwitch.checked ? 'ativada' : 'desativada'}`);
    // });

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
        log('Processando arquivos Excel...');
        
        const dadosExtraidos = [];
        
        // Processar cada arquivo
        for (const arquivo of arquivos) {
            try {
                log(`Processando arquivo: ${arquivo.name}`);
                
                // Verificar se é um arquivo Excel
                if (!arquivo.name.endsWith('.xlsx') && !arquivo.name.endsWith('.xls')) {
                    log(`Arquivo ignorado: ${arquivo.name} não é um arquivo Excel válido`, 'warning');
                    continue;
                }
                
                const arrayBuffer = await arquivo.arrayBuffer();
                const workbook = XLSX.read(arrayBuffer, { type: 'array' });
                
                // Processar a primeira planilha
                const primeiraSheetNome = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[primeiraSheetNome];
                
                // Converter para JSON
                const linhas = XLSX.utils.sheet_to_json(worksheet, { raw: false });
                log(`Encontradas ${linhas.length} linhas em ${arquivo.name}`);
                
                if (linhas.length === 0) {
                    log(`Arquivo vazio ou sem dados tabulares: ${arquivo.name}`, 'warning');
                    continue;
                }
                
                // Analisar cabeçalhos para determinar o formato da planilha
                const primeiraLinha = linhas[0];
                const colunas = Object.keys(primeiraLinha);
                log(`Colunas detectadas: ${colunas.join(', ')}`);
                
                // Identificar colunas importantes
                let colunaId = null;
                let colunaData = null;
                let colunaRecebto = null; // Nova coluna para recebimento
                let colunaValor = null;
                let colunaBPO = null;
                
                // Possíveis nomes para a coluna de ID
                const idColunas = ['Referência do Aplicativo', 'ID', 'ID Aplicativo', 'Código', 'Referencia'];
                for (const col of idColunas) {
                    if (colunas.includes(col)) {
                        colunaId = col;
                        log(`Coluna de ID encontrada: ${colunaId}`);
                        break;
                    }
                }
                
                // Possíveis nomes para a coluna de data de emissão
                const dataColunas = ['Emissão', 'Data', 'Data Pagamento', 'Data Emissão'];
                for (const col of dataColunas) {
                    if (colunas.includes(col)) {
                        colunaData = col;
                        log(`Coluna de data de emissão encontrada: ${colunaData}`);
                        break;
                    }
                }
                
                // Possíveis nomes para a coluna de data de recebimento
                const recbtoColunas = ['Recebto.', 'Recebimento', 'Data Recebimento', 'Data Recebto'];
                for (const col of recbtoColunas) {
                    if (colunas.includes(col)) {
                        colunaRecebto = col;
                        log(`Coluna de data de recebimento encontrada: ${colunaRecebto}`);
                        break;
                    }
                }
                
                // Possíveis nomes para a coluna de valor
                const valorColunas = ['Valor do Repasse', 'Comissão R$', 'Comissão', 'Valor', 'Repasse'];
                for (const col of valorColunas) {
                    if (colunas.includes(col)) {
                        colunaValor = col;
                        log(`Coluna de valor encontrada: ${colunaValor}`);
                        break;
                    }
                }
                
                // Verificar se existe coluna BPO
                const bpoColunas = ['BPO', 'Bpo', 'bpo', 'É BPO'];
                for (const col of bpoColunas) {
                    if (colunas.includes(col)) {
                        colunaBPO = col;
                        log(`Coluna BPO encontrada: ${colunaBPO}`);
                        break;
                    }
                }
                
                // Verificar se encontramos as colunas necessárias
                if (!colunaId) {
                    log(`Não foi possível identificar a coluna de ID no arquivo ${arquivo.name}`, 'error');
                    continue;
                }
                
                if (!colunaData) {
                    log(`Não foi possível identificar a coluna de data no arquivo ${arquivo.name}`, 'warning');
                }
                
                if (!colunaValor) {
                    log(`Não foi possível identificar a coluna de valor no arquivo ${arquivo.name}`, 'error');
                    continue;
                }
                
                // Processar cada linha para extrair informações necessárias
                let linhasProcessadas = 0;
                
                for (const linha of linhas) {
                    if (!linha[colunaId]) continue;
                    
                    const idCompleto = linha[colunaId];
                    const idReferencia = extrairIdReferencia(idCompleto);
                    
                    if (!idReferencia) {
                        continue;
                    }
                    
                    // Extrair o mês de referência, mês de recebimento e valor da comissão
                    let mesReferencia, mesRecebimento, valorComissao;
                    
                    // Função para converter data (número serial do Excel ou string) para objeto Date
                    function excelDateToJSDate(excelDate) {
                        if (!excelDate) return null;
                        
                        // Verificar se é uma string de data (ex: "22/05/2025" ou "2025-05-22")
                        if (typeof excelDate === 'string') {
                            // Tentar formatos comuns no Brasil e internacional
                            let dataParts;
                            
                            // Tentar formato DD/MM/YYYY
                            if (excelDate.includes('/')) {
                                dataParts = excelDate.split('/');
                                if (dataParts.length === 3) {
                                    const day = parseInt(dataParts[0], 10);
                                    const month = parseInt(dataParts[1], 10) - 1; // Meses em JS são 0-indexed
                                    const year = parseInt(dataParts[2], 10);
                                    if (!isNaN(day) && !isNaN(month) && !isNaN(year)) {
                                        return new Date(year, month, day);
                                    }
                                }
                            }
                            
                            // Tentar formato YYYY-MM-DD
                            if (excelDate.includes('-')) {
                                dataParts = excelDate.split('-');
                                if (dataParts.length === 3) {
                                    const year = parseInt(dataParts[0], 10);
                                    const month = parseInt(dataParts[1], 10) - 1; // Meses em JS são 0-indexed
                                    const day = parseInt(dataParts[2], 10);
                                    if (!isNaN(day) && !isNaN(month) && !isNaN(year)) {
                                        return new Date(year, month, day);
                                    }
                                }
                            }
                            
                            // Se nenhum formato conhecido, tentar o construtor Date diretamente
                            const dateObj = new Date(excelDate);
                            if (!isNaN(dateObj.getTime())) {
                                return dateObj;
                            }
                            
                            console.log(`Não foi possível converter a data: ${excelDate}`);
                            return null;
                        }
                        
                        // Se for um número, assumir que é o formato serial do Excel
                        const excelNumber = parseFloat(excelDate);
                        if (!isNaN(excelNumber)) {
                            const EXCEL_EPOCH = new Date(1899, 11, 30); // 30/12/1899
                            const millisecondsPerDay = 24 * 60 * 60 * 1000;
                            return new Date(EXCEL_EPOCH.getTime() + excelNumber * millisecondsPerDay);
                        }
                        
                        return null;
                    }
                    
                    // Obter a data de emissão da coluna identificada
                    const dataEmissao = colunaData ? excelDateToJSDate(linha[colunaData]) : null;
                    
                    // Obter a data de recebimento da coluna identificada (se existir)
                    const dataRecebimento = colunaRecebto ? excelDateToJSDate(linha[colunaRecebto]) : null;
                    
                    // Se tiver data de recebimento, usar ela como referência, senão usar data de emissão
                    mesReferencia = dataRecebimento || dataEmissao || new Date(); // Usar data atual se não encontrar
                    
                    // Obter o valor da coluna identificada
                    valorComissao = parseFloat(linha[colunaValor]) || 0;
                    
                    // Determinar se é um registro BPO
                    let isBPO = filesData.isBPO; // Valor padrão do switch button
                    
                    // Se existe coluna BPO no arquivo, usar esse valor com precedência
                    if (colunaBPO && linha[colunaBPO] !== undefined) {
                        const valorBPO = linha[colunaBPO].toString().trim().toLowerCase();
                        // Verificar se o valor é 'sim' ou equivalente
                        if (valorBPO === 'sim' || valorBPO === 's' || valorBPO === 'true' || valorBPO === '1' || valorBPO === 'yes' || valorBPO === 'y') {
                            isBPO = true;
                        } else if (valorBPO === 'não' || valorBPO === 'nao' || valorBPO === 'n' || valorBPO === 'false' || valorBPO === '0' || valorBPO === 'no') {
                            isBPO = false;
                        }
                        // Em caso de valor não reconhecido, manter o padrão do switch
                    }
                    
                    // Formatar o mês como string (ex: "Jan", "Fev", etc.)
                    const meses = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'];
                    const mesFormatado = mesReferencia ? meses[mesReferencia.getMonth()] : 'N/A';
                    
                    // Verificar se os meses de emissão e recebimento são diferentes
                    let mesesDiferentes = false;
                    let mensagemDiferenca = '';
                    
                    if (dataEmissao && dataRecebimento) {
                        // Verificar se mês ou ano são diferentes
                        mesesDiferentes = dataEmissao.getMonth() !== dataRecebimento.getMonth() || 
                                         dataEmissao.getFullYear() !== dataRecebimento.getFullYear();
                        
                        // Se forem diferentes, criar uma mensagem detalhada
                        if (mesesDiferentes) {
                            // Obter nomes dos meses
                            const mesesNomes = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'];
                            const mesEmissao = mesesNomes[dataEmissao.getMonth()];
                            const mesRecebimento = mesesNomes[dataRecebimento.getMonth()];
                            const anoEmissao = dataEmissao.getFullYear();
                            const anoRecebimento = dataRecebimento.getFullYear();
                            
                            // Formatar a mensagem
                            if (anoEmissao !== anoRecebimento) {
                                mensagemDiferenca = `O mês de recebimento foi em ${mesRecebimento}/${anoRecebimento}, mas a data de emissão é de ${mesEmissao}/${anoEmissao}.`;
                            } else {
                                mensagemDiferenca = `O mês de recebimento foi em ${mesRecebimento}, mas a data de emissão é do mês de ${mesEmissao}.`;
                            }
                        }
                    }
                    
                    dadosExtraidos.push({
                        idReferencia,
                        idCompleto,
                        mesReferencia,
                        mesFormatado,
                        valorComissao: parseFloat(valorComissao),
                        fonte: arquivo.name,
                        isBPO: isBPO, // Adicionar a flag de BPO aos dados
                        mesesDiferentes, // Flag para indicar se os meses de emissão e recebimento são diferentes
                        mensagemDiferenca, // Mensagem detalhada sobre a diferença de meses
                        linhaOriginal: linha
                    });
                    
                    linhasProcessadas++;
                }
                
                log(`Processadas ${linhasProcessadas} linhas válidas do arquivo ${arquivo.name}`);
                
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
            
            // Primeiro lemos com XLSX para obter os valores já calculados das fórmulas
            log('Lendo valores calculados com SheetJS...');
            const workbookXLSX = XLSX.read(arrayBuffer, { type: 'array' });
            const sheetName = "APPs";
            const sheetXLSX = workbookXLSX.Sheets[sheetName];
            
            if (!sheetXLSX) {
                throw new Error(`Planilha "${sheetName}" não encontrada na planilha de controle.`);
            }
            
            // Converter a planilha para um array de objetos com os valores calculados
            const valoresCalculados = XLSX.utils.sheet_to_json(sheetXLSX, { header: 1, raw: false });
            log(`Leitura concluída. Foram encontradas ${valoresCalculados.length} linhas.`);
            
            // Agora carregamos com ExcelJS para manter formatação e editar
            log('Carregando planilha com ExcelJS para preservar formatação...');
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.load(arrayBuffer);
            
            const worksheet = workbook.getWorksheet(sheetName);
            if (!worksheet) {
                throw new Error(`Planilha "${sheetName}" não encontrada na planilha de controle (ExcelJS).`);
            }
            
            // Obter os cabeçalhos da planilha Excel (para preservar o comportamento original)
            const headerRow = worksheet.getRow(1);
            const headers = [];
            const colunasPorNome = {};
            
            headerRow.eachCell({ includeEmpty: false }, (cell, colNumber) => {
                const cellValue = cell.value ? cell.value.toString() : '';
                headers[colNumber] = cellValue;
                if (cellValue) {
                    colunasPorNome[cellValue] = colNumber;
                }
            });
            
            // Criar um mapeamento para valores calculados
            const xlsxHeaders = valoresCalculados[0] || [];
            const colunaXlsxParaExcelJS = {};
            
            xlsxHeaders.forEach((header, xlsxIndex) => {
                if (header) {
                    // Encontrar o índice correspondente em ExcelJS
                    Object.keys(colunasPorNome).forEach(key => {
                        if (key.toString() === header.toString()) {
                            colunaXlsxParaExcelJS[xlsxIndex] = colunasPorNome[key];
                        }
                    });
                }
            });
            
            log(`Mapeamento de colunas concluído. Encontradas ${Object.keys(colunasPorNome).length} colunas.`);
            
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
                const { idReferencia, mesFormatado, valorComissao, isBPO } = registro;
                
                // Verificar se o registro segue regras de BPO
                log(`Processando registro ${idReferencia}: ${isBPO ? 'BPO' : 'Não BPO'}`);
                
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
                log(`Procurando coluna ${colunaMes} para mês ${mesFormatado}. Encontrada: ${colunaIndex ? 'Sim' : 'Não'}`);
                
                // Se não encontrou pelo nome exato, tentar nomes alternativos
                if (!colunaIndex) {
                    // Mapeamento de nomes alternativos para meses
                    const nomesAlternativos = {
                        'Jan': ['janeiro', 'jan'],
                        'Fev': ['fevereiro', 'fev'],
                        'Mar': ['março', 'mar'],
                        'Abr': ['abril', 'abr'],
                        'Mai': ['maio', 'mai'],
                        'Jun': ['junho', 'jun'],
                        'Jul': ['julho', 'jul'],
                        'Ago': ['agosto', 'ago'],
                        'Set': ['setembro', 'set'],
                        'Out': ['outubro', 'out'],
                        'Nov': ['novembro', 'nov'],
                        'Dez': ['dezembro', 'dez']
                    };
                    
                    // Obter alternativas para o mês atual
                    const alternativas = nomesAlternativos[mesFormatado] || [];
                    log(`Buscando alternativas para ${mesFormatado}: ${alternativas.join(', ')}`);
                    
                    // Procurar colunas que contenham qualquer uma das alternativas
                    const colunasRelacionadas = Object.keys(colunasPorNome).filter(coluna => {
                        const colunaLower = coluna.toLowerCase();
                        return colunaLower.includes(mesFormatado.toLowerCase()) || 
                               alternativas.some(alt => colunaLower.includes(alt));
                    });
                    
                    log(`Colunas relacionadas encontradas: ${colunasRelacionadas.length}`);
                    
                    if (colunasRelacionadas.length > 0) {
                        // Mostrar todas as colunas encontradas para debug
                        colunasRelacionadas.forEach(col => {
                            log(`Coluna alternativa candidata para ${mesFormatado}: ${col}`);
                        });
                        
                        // Usar a primeira coluna alternativa encontrada
                        log(`Usando coluna alternativa para ${mesFormatado}: ${colunasRelacionadas[0]}`);
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
                    
                    // Buscar o ID nos dados obtidos via SheetJS (que já tem os valores calculados)
                    log(`Buscando ID ${idReferencia} nos dados do SheetJS...`);
                    
                    // Primeiro, encontrar o índice da coluna no SheetJS
                    const xlsxHeaderRow = valoresCalculados[0] || [];
                    let xlsxColIndex = -1;
                    
                    xlsxHeaderRow.forEach((header, index) => {
                        if (header && header.toString() === colunaId) {
                            xlsxColIndex = index;
                        }
                    });
                    
                    if (xlsxColIndex === -1) {
                        log(`Aviso: Não foi possível encontrar a coluna ${colunaId} nos dados do SheetJS`, 'warning');
                    } else {
                        log(`Coluna ${colunaId} encontrada no índice ${xlsxColIndex} do SheetJS`);
                        
                        for (let i = 1; i < valoresCalculados.length; i++) { // Começamos do 1 para pular o cabeçalho
                            const linha = valoresCalculados[i];
                            if (!linha) continue;
                            
                            const valorCelulaCalc = linha[xlsxColIndex];
                            
                            if (valorCelulaCalc && valorCelulaCalc.toString().includes(idReferencia)) {
                                log(`ID ${idReferencia} encontrado na linha ${i + 1} (via SheetJS)`); // +1 porque i é 0-indexed mas as linhas na UI são 1-indexed
                                linhaEncontrada = i + 1; // +1 para ajustar para a linha real no ExcelJS
                                celulaEncontrada = true;
                                break;
                            }
                        }
                    }
                        
                    if (!celulaEncontrada) {
                        // Fallback para o método antigo se não encontrar
                        let linhaAtual = 2; // Começar após o cabeçalho
                        
                        while (!celulaEncontrada && linhaAtual <= worksheet.rowCount) {
                            const cell = worksheet.getRow(linhaAtual).getCell(colunaIdIndex);
                            let valorCelula = '';
                            
                            if (cell.result !== undefined) {
                                valorCelula = cell.result;
                            } else if (cell.value && typeof cell.value === 'object' && cell.value.result !== undefined) {
                                valorCelula = cell.value.result;
                            } else if (cell.text) {
                                valorCelula = cell.text;
                            } else if (cell.value) {
                                valorCelula = cell.value.toString();
                            }
                            
                            if (valorCelula && valorCelula.includes(idReferencia)) {
                                log(`ID ${idReferencia} encontrado na linha ${linhaAtual} (via fallback)`); 
                                linhaEncontrada = linhaAtual;
                                celulaEncontrada = true;
                                break;
                            }
                            linhaAtual++;
                        }
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
                
                // Verificar se devemos atualizar esta célula com base nas regras de BPO
                let deveAtualizar = true;
                
                // Verificar a coluna "É BPO" na planilha controle (se existir)
                const colunaBPOIndex = colunasPorNome['É BPO'] || colunasPorNome['BPO'];
                let bpoNaPlanilhaControle = false;
                
                if (colunaBPOIndex) {
                    const cellBPO = worksheet.getRow(linhaEncontrada).getCell(colunaBPOIndex);
                    const valorBPO = cellBPO.value;
                    
                    if (valorBPO) {
                        const strBPO = valorBPO.toString().toLowerCase();
                        bpoNaPlanilhaControle = (strBPO === 'sim' || strBPO === 's' || strBPO === 'true' || strBPO === '1');
                        log(`Valor BPO na planilha de controle para ${idReferencia}: ${bpoNaPlanilhaControle ? 'BPO' : 'Não BPO'}`);
                    }
                }
                
                // Regras de atualização baseadas no status de BPO:
                // Se o registro é BPO mas a planilha controle não é BPO, não atualizar
                // Se o registro não é BPO mas a planilha controle é BPO, não atualizar
                if (colunaBPOIndex && isBPO !== bpoNaPlanilhaControle) {
                    log(`Ignorando atualização para ${idReferencia} - ${mesFormatado} - Status BPO incompatível`, 'warning');
                    deveAtualizar = false;
                }
                
                // Verificar melhor se a célula está vazia
                const celulaEstaVazia = valorAtual === null || valorAtual === undefined || 
                                        valorAtual === '' || 
                                        (typeof valorAtual === 'object' && (!valorAtual.result || valorAtual.result === ''));
                
                // Log para debug do valor atual
                log(`Valor atual para ${idReferencia} - ${mesFormatado}: ${celulaEstaVazia ? 'VAZIO' : valorAtualNumerico} (coluna: ${colunaMes})`);
                
                // Se o valor atual for diferente do novo valor e devemos atualizar, atualizar
                if (deveAtualizar && (celulaEstaVazia || Math.abs(valorAtualNumerico - valorComissao) > 0.01)) {
                    log(`Atualizando ${idReferencia} - ${mesFormatado}: ${valorAtual || 'VAZIO'} -> ${valorComissao} ${isBPO ? '[BPO]' : ''}`);
                    
                    // Atualizar o valor na célula
                    cell.value = valorComissao;
                    
                    // Verificar se os meses de emissão e recebimento são diferentes
                    if (registro.mesesDiferentes) {
                        const cellRef = `${worksheet.getColumn(colunaIndex).letter}${linhaEncontrada}`;
                        log(`Meses diferentes detectados para ${idReferencia}, adicionando nota à célula ${cellRef}`);
                        
                        // Obter a mensagem detalhada sobre a diferença de meses
                        const mensagem = registro.mensagemDiferenca || 'ATENÇÃO: Os meses de emissão e recebimento são diferentes.';
                        
                        // Adicionar um comentário à célula explicando a diferença de meses
                        if (worksheet.getCell(cellRef).comment) {
                            // Se já existe um comentário, atualizá-lo
                            worksheet.getCell(cellRef).comment.texts = [
                                {'font': {'size': 12}, 'text': mensagem}
                            ];
                        } else {
                            // Se não existe um comentário, criar um novo
                            worksheet.getCell(cellRef).note = mensagem;
                        }
                        
                        log(`Célula ${cellRef} com comentário adicionado: "${mensagem}"`);
                    }
                    
                    // Registrar a célula alterada para o relatório
                    celulasAlteradas.push({
                        idReferencia,
                        mesReferencia: mesFormatado,
                        colunaMes,
                        valorAntigo: valorAtual || 'VAZIO',
                        valorNovo: valorComissao,
                        isBPO: isBPO,
                        mesesDiferentes: registro.mesesDiferentes,
                        celula: `${worksheet.getColumn(colunaIndex).letter}${linhaEncontrada}`
                    });
                    
                    atualizacoes++;
                } else if (!deveAtualizar) {
                    log(`Não atualizado ${idReferencia} - ${mesFormatado}: Regras de BPO não permitem`);
                } else {
                    log(`Valor igual ao existente para ${idReferencia} - ${mesFormatado}: ${valorComissao}`);
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
                    relatorioConteudo += `${index + 1}. ID: ${celula.idReferencia} - Mês: ${celula.mesReferencia} (Célula: ${celula.celula})${celula.mesesDiferentes ? ' [MESES DIFERENTES]' : ''}\n`;
                    relatorioConteudo += `   Valor anterior: ${celula.valorAntigo.result}\n`;
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
                item.valorComissao > 0 &&
                !item.isBPO
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
