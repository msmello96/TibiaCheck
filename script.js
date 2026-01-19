async function processarArquivo() {
    const fileInput = document.getElementById('excelFile');
    const file = fileInput.files[0];
    if (!file) {
        alert('Selecione um arquivo Excel!');
        return;
    }

    // Ler o arquivo Excel
    const reader = new FileReader();
    reader.onload = async function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        // Encontrar o índice da coluna 'Nick do Character'
        const colunaIndex = jsonData[0].indexOf('Nick do Character');
        if (colunaIndex === -1) {
            alert('Coluna "Nick do Character" não encontrada!');
            return;
        }

        // Coletar todos os valores da coluna
        const valores = jsonData.slice(1)
            .map(row => row[colunaIndex])
            .filter(valor => valor !== undefined && valor !== null && typeof valor === 'string' && valor.trim() !== '');

        // Mostrar progresso inicial
        const divResultado = document.getElementById('resultado');
        divResultado.innerHTML = `<p>Processando ${valores.length} personagens... Por favor, aguarde.</p>`;

        // Processar com delay entre requisições
        const resultados = [];
        let processados = 0;
        
        // Pegar o delay configurado pelo usuário
        const delayConfig = parseInt(document.getElementById('delayConfig').value) || 3000;

        for (const valor of valores) {
            // Delay ANTES da requisição (exceto na primeira)
            if (processados > 0) {
                await sleep(delayConfig);
            }
            try {
                // Fazer requisição com retry
                const data = await fetchComRetry(valor.trim());
                
                if (data && data.character && data.character.character) {
                    const char = data.character.character;
                    const name = char.name;
                    const guild = (char.guild && char.guild.name) ? char.guild.name : 'Sem Guild';
                    resultados.push({ name, guild });
                } else {
                    resultados.push({ name: valor, guild: 'Não encontrado' });
                }
            } catch (error) {
                console.error(`Erro ao processar ${valor}: ${error}`);
                resultados.push({ name: valor, guild: 'Erro na consulta' });
            }

            processados++;
            // Atualizar progresso
            divResultado.innerHTML = `<p>Processando: ${processados}/${valores.length} personagens...</p>`;
        }

        // Exibir resultados finais
        exibirResultados(resultados);
    };
    reader.readAsArrayBuffer(file);
}

// Função para fazer requisição com retry
async function fetchComRetry(characterName, tentativas = 3) {
    const apiUrl = `https://api.tibiadata.com/v4/character/${encodeURIComponent(characterName)}`;
    
    for (let i = 0; i < tentativas; i++) {
        try {
            const response = await fetch(apiUrl);
            
            // Se receber 503, esperar mais tempo antes de tentar novamente
            if (response.status === 503) {
                if (i < tentativas - 1) {
                    const waitTime = (i + 1) * 5000; // 5s, 10s, 15s
                    console.log(`503 recebido para ${characterName}. Aguardando ${waitTime/1000}s antes de tentar novamente...`);
                    await sleep(waitTime);
                    continue;
                }
                throw new Error('Serviço indisponível após múltiplas tentativas');
            }

            if (!response.ok) {
                throw new Error(`HTTP ${response.status}`);
            }

            return await response.json();
        } catch (error) {
            console.log(`Erro ao buscar ${characterName} (tentativa ${i + 1}/${tentativas}):`, error.message);
            if (i === tentativas - 1) {
                throw error;
            }
            // Esperar antes de tentar novamente
            await sleep(35000); // 35 segundos
        }
    }
}

// Função auxiliar para criar delay
function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

function exibirResultados(resultados) {
    const divResultado = document.getElementById('resultado');
    if (resultados.length === 0) {
        divResultado.innerHTML = '<p>Nenhum resultado encontrado.</p>';
        return;
    }

    let counter = 0;
    let tabela = '<table><tr><th>ID</th><th>Character Name</th><th>Guild</th></tr>';
    resultados.forEach(item => {
        counter += 1;
        tabela += `<tr><td>${counter}</td><td>${item.name}</td><td>${item.guild}</td></tr>`;
    });
    tabela += '</table>';
    
    // Adicionar resumo detalhado
    const hellsscreamReturns = resultados.filter(r => r.guild === 'Hellsscream Returns').length;
    const outrasGuilds = resultados.filter(r => 
        r.guild !== 'Hellsscream Returns' && 
        r.guild !== 'Sem Guild' && 
        r.guild !== 'Não encontrado' && 
        r.guild !== 'Erro na consulta'
    ).length;
    const semGuild = resultados.filter(r => r.guild === 'Sem Guild').length;
    const erros = resultados.filter(r => r.guild === 'Não encontrado' || r.guild === 'Erro na consulta').length;
    
    let resumo = `<p><strong>Resumo:</strong> ${resultados.length} personagens processados | `;
    resumo += `${hellsscreamReturns} em Hellsscream Returns | `;
    resumo += `${outrasGuilds} em outras guilds | `;
    resumo += `${semGuild} sem guild`;
    if (erros > 0) {
        resumo += ` | ${erros} com erro`;
    }
    resumo += `</p>`;
    
    divResultado.innerHTML = resumo + tabela;
}
