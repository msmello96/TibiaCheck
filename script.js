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
        const sheetName = workbook.SheetNames[0]; // Pega a primeira aba
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        // Encontrar o índice da coluna 'QUAL SEU NICK'
        const colunaIndex = jsonData[0].indexOf('Nick do Character');
        if (colunaIndex === -1) {
            alert('Coluna Nick do Character" não encontrada!');
            return;
        }

        // Coletar todos os valores da coluna (ignorando cabeçalho e valores vazios)
        const valores = jsonData.slice(1)
            .map(row => row[colunaIndex])
            .filter(valor => valor !== undefined && valor !== null && typeof valor === 'string' && valor.trim() !== '');

        // Para cada valor, chamar a API e coletar dados
        const resultados = [];
        for (const valor of valores) {
            try {
                const response = await fetch(`https://api.tibiadata.com/v4/character/${encodeURIComponent(valor.trim())}`);
                if (!response.ok) throw new Error('Erro na API');
                const json = await response.json();
                const char = json.character.character;
                const name = char.name;
                const guild = (char.guild && char.guild.name) ? char.guild.name : 'Sem Guild';
                resultados.push({ name, guild });
            } catch (error) {
                console.error(`Erro ao processar ${valor}: ${error}`);
                // Opcional: adicionar um resultado de erro, ex: resultados.push({ name: valor, guild: 'Erro na API' });
            }
        }

        // Exibir resultados em uma tabela
        exibirResultados(resultados);
    };
    reader.readAsArrayBuffer(file);
}

function exibirResultados(resultados) {
    const divResultado = document.getElementById('resultado');
    if (resultados.length === 0) {
        divResultado.innerHTML = '<p>Nenhum resultado encontrado.</p>';
        return;
    }
    let counter = 0;
    let tabela = '<table><th>ID</th><tr><th>Character Name</th><th>Guild</th></tr>';
    resultados.forEach(item => {
        counter += 1;
        tabela += `<tr><td>${counter}</td><td>${item.name}</td><td>${item.guild}</td></tr>`;
    });
    tabela += '</table>';
    divResultado.innerHTML = tabela;
}
