const { exec } = require("child_process");

function gerarExecutavelPython() {
    const comando = 'pyinstaller --onefile "SOMA.py"';

    exec(comando, (erro, stdout, stderr) => {
        if (erro) {
            console.error(`Erro ao executar o comando: ${erro.message}`);
            return;
        }

        // Exibir stderr apenas se não contiver logs genéricos do PyInstaller
        if (stderr && !stderr.includes("INFO")) {
            console.error(`Erro no processo: ${stderr}`);
        }

        console.log(`Saída do comando:\n${stdout}`);
    });
}

// Executa a função
gerarExecutavelPython();
