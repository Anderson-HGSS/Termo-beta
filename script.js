const linhas = document.querySelectorAll("#respostas > div");

const teclado = document.querySelector("#teclado");

const botoes = teclado.querySelectorAll("button");



const botaoEnter = document.querySelector("#enter");

const botaoBackspace = document.querySelector("#backspace");
const entradaMobile = document.querySelector("#entrada-mobile");

// Mantém um campo de texto focável para abrir o teclado nativo em celulares.
function focarEntradaMobile() {
    if (entradaMobile && !jogoFinalizado) {
        entradaMobile.focus({ preventScroll: true });
    }
}



// Palavras aceitas como tentativa (lista ampla).

let palavrasDicionario = new Set();



// Palavras que podem ser sorteadas (lista curada, mais comuns).

let palavrasSecretas = [];



let palavraSecreta = "";



let tentativaAtual = 0;

let letraAtual = 0;

let jogoFinalizado = false;



// Controle dos elementos injetados via JavaScript.

let modalFimDeJogo = null;

let temporizadorModal = null;



let avisoAtual = null;

let temporizadoresAviso = [];

let animacaoShake = null;



// Normaliza palavras para permitir comparações sem acentos.

function normalizarPalavra(palavra) {

    return palavra

        .normalize("NFD")

        .replace(/[\u0300-\u036f]/g, "")

        .toUpperCase();

}



// Carrega a biblioteca e prepara o dicionário.

async function carregarDicionario() {

    try {

        console.log("Carregando dicionário...");



        const modulo = await import(

            "https://esm.sh/@andsfonseca/palavras-pt-br"

        );



        const { Word, UNVERSEDV2, BRISPELL, PYTHONPROBR } = modulo;



        if (!Word) {

            throw new Error("Não foi possível acessar a biblioteca.");

        }



        // Extrai de uma base as palavras sem acentos, compostos,

        // caracteres especiais ou nomes próprios, mantendo

        // somente as que têm exatamente cinco letras.

        const extrairPalavras = base => {

            if (base) {

                Word.library = base;

            }



            return Word.getAllWords(5, true, false, false, false)

                .map(normalizarPalavra)

                .filter(palavra => /^[A-Z]{5}$/.test(palavra));

        };



        // Lista curada (padrão da biblioteca): usada para sortear,

        // pois reúne palavras mais comuns e jogáveis.

        const secretas = new Set(extrairPalavras(UNVERSEDV2));

        palavrasSecretas = [...secretas];



        // Lista ampla: une as demais bases para aceitar mais tentativas.

        palavrasDicionario = new Set(secretas);



        [BRISPELL, PYTHONPROBR].forEach(base => {

            if (base) {

                extrairPalavras(base).forEach(palavra =>

                    palavrasDicionario.add(palavra)

                );

            }

        });



        if (palavrasDicionario.size === 0) {

            throw new Error("O dicionário não retornou palavras válidas.");

        }



        sortearPalavra();



        inicializarJogo();



        console.log(

            "Dicionário carregado:",

            palavrasDicionario.size,

            "palavras aceitas,",

            palavrasSecretas.length,

            "sorteáveis."

        );



    } catch (erro) {

        console.error("Erro ao carregar o dicionário:", erro);



        alert(

            "Não foi possível carregar o dicionário. " +

            "Verifique sua conexão com a internet."

        );

    }

}



// Sorteia uma palavra válida do dicionário.

function sortearPalavra() {

    const indice = Math.floor(

        Math.random() * palavrasSecretas.length

    );



    palavraSecreta = palavrasSecretas[indice];



    // console.log("Palavra sorteada:", palavraSecreta);

}



// Retorna as casas de uma linha.

function obterLetrasDaLinha(indice) {

    return linhas[indice].querySelectorAll(".letra");

}



// Inicializa o tabuleiro.

function inicializarJogo() {

    linhas.forEach((linha, indice) => {

        linha.classList.toggle(

            "resposta-atual",

            indice === 0

        );



        const letras = obterLetrasDaLinha(indice);



        letras.forEach(letra => {

            letra.textContent = "";

            letra.className = "letra letra-vazia";

        });

    });



    atualizarTeclado();

    atualizarLinhaAtual();

}



// Atualiza a casa destacada na tentativa atual.

function atualizarLinhaAtual() {

    linhas.forEach((linha, indice) => {

        linha.classList.toggle(

            "resposta-atual",

            indice === tentativaAtual && !jogoFinalizado

        );



        const letras = obterLetrasDaLinha(indice);



        letras.forEach((letra, posicao) => {

            if (indice !== tentativaAtual || jogoFinalizado) {

                return;

            }



            if (posicao === letraAtual && letraAtual < 5) {

                letra.classList.remove(

                    "letra-vazia",

                    "letra-atual"

                );



                letra.classList.add("letra-atual");

            } else if (letra.textContent === "") {

                letra.className = "letra letra-vazia";

            } else {

                letra.classList.remove("letra-atual");

            }

        });

    });

}



// Insere uma letra.

function inserirLetra(letra) {

    if (jogoFinalizado) {

        return;

    }



    const letras = obterLetrasDaLinha(tentativaAtual);



    // Escreve na casa selecionada (substitui a letra, se já houver uma).

    letras[letraAtual].textContent =

        normalizarPalavra(letra);



    // Avança para a próxima casa; na última, o cursor permanece nela.

    if (letraAtual < 4) {

        letraAtual++;

    }



    atualizarLinhaAtual();

}



// Move a casa selecionada com as setas (-1 = esquerda, 1 = direita).

function moverCursor(direcao) {

    if (jogoFinalizado) {

        return;

    }



    letraAtual = Math.min(4, Math.max(0, letraAtual + direcao));



    atualizarLinhaAtual();

}



// Apaga a última letra digitada.

function apagarLetra() {

    if (jogoFinalizado) {

        return;

    }



    const letras = obterLetrasDaLinha(tentativaAtual);



    // Se a casa selecionada está vazia, volta uma casa antes de apagar.

    if (letras[letraAtual].textContent === "") {

        if (letraAtual === 0) {

            return;

        }



        letraAtual--;

    }



    letras[letraAtual].textContent = "";

    letras[letraAtual].className = "letra letra-vazia";



    atualizarLinhaAtual();

}



/* ========================================================*

 *   FEEDBACK DE PALAVRA NÃO ENCONTRADA (SHAKE + LABEL)

 *   ======================================================== */



const DURACAO_SHAKE_MS = 500;      // 0.5 segundo de tremor

const DURACAO_AVISO_MS = 2000;     // 2 segundos visível

const DURACAO_FADE_MS = 500;       // duração do fade-out



// Remove imediatamente qualquer aviso ativo e seus temporizadores.

function removerAviso() {

    temporizadoresAviso.forEach(clearTimeout);

    temporizadoresAviso = [];



    if (avisoAtual) {

        avisoAtual.getAnimations().forEach(animacao => animacao.cancel());

        avisoAtual.remove();

        avisoAtual = null;

    }

}



// Aplica o tremor horizontal na linha atual.

function tremerLinha(linha) {

    if (animacaoShake) {

        animacaoShake.cancel();

    }



    animacaoShake = linha.animate(

        [

            { transform: "translateX(0)" },

            { transform: "translateX(-12px)" },

            { transform: "translateX(12px)" },

            { transform: "translateX(-10px)" },

            { transform: "translateX(10px)" },

            { transform: "translateX(-6px)" },

            { transform: "translateX(6px)" },

            { transform: "translateX(0)" }

        ],

        {

            duration: DURACAO_SHAKE_MS,

            easing: "linear"

        }

    );

}



// Exibe a label ao lado da linha atual.

function exibirAviso(linha, texto) {

    removerAviso();



    const letras = linha.querySelectorAll(".letra");

    const primeira = letras[0].getBoundingClientRect();

    const ultima = letras[letras.length - 1].getBoundingClientRect();



    const aviso = document.createElement("div");

    aviso.textContent = texto;



    Object.assign(aviso.style, {

        position: "absolute",

        zIndex: "900",

        color: "#ffffff",

        backgroundColor: "#3b3b3b",

        border: "2px solid #b3b3b3",

        borderRadius: "10px",

        padding: "10px 16px",

        fontSize: "16px",

        fontWeight: "bold",

        whiteSpace: "nowrap",

        pointerEvents: "none",

        opacity: "1"

    });



    document.body.appendChild(aviso);



    const margem = 15;

    const larguraAviso = aviso.offsetWidth;

    const alturaAviso = aviso.offsetHeight;



    // Preferência: à direita da linha. Se não couber, vai para a esquerda.

    let esquerda = ultima.right + margem;



    if (esquerda + larguraAviso > window.innerWidth) {

        esquerda = primeira.left - margem - larguraAviso;

    }



    esquerda = Math.max(8, esquerda);



    const topo =

        primeira.top + (primeira.height - alturaAviso) / 2;



    aviso.style.left = (esquerda + window.scrollX) + "px";

    aviso.style.top = (topo + window.scrollY) + "px";



    avisoAtual = aviso;



    // Após 2 segundos visível, inicia o fade-out e remove do DOM.

    const temporizadorFade = setTimeout(() => {

        if (avisoAtual !== aviso) {

            return;

        }



        const fade = aviso.animate(

            [{ opacity: 1 }, { opacity: 0 }],

            {

                duration: DURACAO_FADE_MS,

                easing: "ease-out",

                fill: "forwards"

            }

        );



        fade.onfinish = () => {

            aviso.remove();



            if (avisoAtual === aviso) {

                avisoAtual = null;

            }

        };

    }, DURACAO_AVISO_MS);



    temporizadoresAviso.push(temporizadorFade);

}



// Efeito completo para tentativa inválida (shake + label temporária).

function feedbackPalavraInvalida(texto = "Palavra não encontrada") {

    const linha = linhas[tentativaAtual];



    exibirAviso(linha, texto);

    tremerLinha(linha);

}



/* ========================================================*

 *   MODAL DE VITÓRIA

 *   ======================================================== */



// Cria e exibe o modal de fim de jogo (vitória ou derrota) no DOM.

function exibirModalFimDeJogo(venceu) {

    fecharModal();



    const corTema = venceu ? "#13bdb4" : "#d63a3a";



    const sobreposicao = document.createElement("div");



    Object.assign(sobreposicao.style, {

        position: "fixed",

        top: "0",

        left: "0",

        width: "100%",

        height: "100%",

        backgroundColor: "rgba(0, 0, 0, 0.8)",

        display: "flex",

        justifyContent: "center",

        alignItems: "center",

        zIndex: "1000"

    });



    const caixa = document.createElement("div");



    Object.assign(caixa.style, {

        backgroundColor: "rgb(31, 31, 31)",

        color: "azure",

        border: "3px solid " + corTema,

        borderRadius: "20px",

        boxShadow: corTema + " 0px 0px 20px",

        padding: "40px 50px",

        textAlign: "center",

        display: "flex",

        flexDirection: "column",

        alignItems: "center",

        gap: "15px",

        maxWidth: "90%"

    });



    const titulo = document.createElement("h2");

    titulo.textContent = venceu ? "Parabéns!" : "Fim de jogo!";

    titulo.style.fontSize = "32px";



    const mensagem = document.createElement("p");

    mensagem.textContent = venceu

        ? "Você acertou a palavra!"

        : "Não foi dessa vez. A palavra era:";

    mensagem.style.fontSize = "18px";



    const palavra = document.createElement("div");

    palavra.textContent = palavraSecreta;



    Object.assign(palavra.style, {

        backgroundColor: corTema,

        color: "white",

        borderRadius: "10px",

        padding: "8px 20px",

        fontSize: "28px",

        fontWeight: "bold",

        letterSpacing: "6px"

    });



    const botaoReiniciar = document.createElement("button");

    botaoReiniciar.textContent = "JOGAR NOVAMENTE";



    Object.assign(botaoReiniciar.style, {

        marginTop: "10px",

        padding: "12px 24px",

        fontSize: "18px",

        color: "white",

        backgroundColor: "#3b3b3b",

        border: "#3b3b3b solid 1px",

        borderRadius: "10px",

        cursor: "pointer",

        transition: "0.3s"

    });



    botaoReiniciar.addEventListener("mouseenter", () => {

        botaoReiniciar.style.boxShadow = "#acacac 0px 3px";

        botaoReiniciar.style.transform = "translate(0px, -3px)";

    });



    botaoReiniciar.addEventListener("mouseleave", () => {

        botaoReiniciar.style.boxShadow = "none";

        botaoReiniciar.style.transform = "none";

    });



    botaoReiniciar.addEventListener("click", reiniciarJogo);



    caixa.append(titulo, mensagem, palavra, botaoReiniciar);

    sobreposicao.appendChild(caixa);

    document.body.appendChild(sobreposicao);



    modalFimDeJogo = sobreposicao;



    // Entrada suave do modal.

    sobreposicao.animate(

        [{ opacity: 0 }, { opacity: 1 }],

        { duration: 300, easing: "ease-out" }

    );



    caixa.animate(

        [

            { transform: "scale(0.9)" },

            { transform: "scale(1)" }

        ],

        { duration: 300, easing: "ease-out" }

    );



    botaoReiniciar.focus();

}



// Remove o modal do DOM, se existir.

function fecharModal() {

    clearTimeout(temporizadorModal);

    temporizadorModal = null;



    if (modalFimDeJogo) {

        modalFimDeJogo.remove();

        modalFimDeJogo = null;

    }

}



// Reinicia a partida: limpa o tabuleiro e sorteia nova palavra.

function reiniciarJogo() {

    fecharModal();

    removerAviso();



    tentativaAtual = 0;

    letraAtual = 0;

    jogoFinalizado = false;



    sortearPalavra();

    inicializarJogo();

}



/* ======================================================== */



// Verifica a palavra digitada.

function verificarPalavra() {

    if (jogoFinalizado) {

        return;

    }



    const letras = obterLetrasDaLinha(tentativaAtual);



    // Com a navegação por setas, a palavra só está completa

    // quando todas as cinco casas estão preenchidas.

    const palavraIncompleta = Array.from(letras).some(

        letra => letra.textContent === ""

    );



    if (palavraIncompleta) {

        feedbackPalavraInvalida("A palavra precisa ter 5 letras");

        return;

    }



    const palavraDigitada = Array.from(letras)

        .map(letra => letra.textContent)

        .join("");



    // Consulta o dicionário em vez de uma lista manual.

    if (!palavrasDicionario.has(palavraDigitada)) {

        feedbackPalavraInvalida();

        return;

    }



    const resultado = Array(5).fill("letra-incorreta");



    const letrasSecretas = palavraSecreta.split("");

    const letrasDigitadas = palavraDigitada.split("");



    // Primeira passagem: letras nas posições corretas.

    for (let i = 0; i < 5; i++) {

        if (letrasDigitadas[i] === letrasSecretas[i]) {

            resultado[i] = "letra-correta";



            letrasSecretas[i] = null;

            letrasDigitadas[i] = null;

        }

    }



    // Segunda passagem: letras presentes em outras posições.

    for (let i = 0; i < 5; i++) {

        if (letrasDigitadas[i] === null) {

            continue;

        }



        const indiceEncontrado =

            letrasSecretas.indexOf(letrasDigitadas[i]);



        if (indiceEncontrado !== -1) {

            resultado[i] = "letra-contem";

            letrasSecretas[indiceEncontrado] = null;

        }

    }



    // Aplica as cores no tabuleiro e no teclado.

    for (let i = 0; i < 5; i++) {

        letras[i].className = "letra " + resultado[i];



        atualizarCorTecla(

            palavraDigitada[i],

            resultado[i]

        );

    }



    // Vitória.

    if (palavraDigitada === palavraSecreta) {

        jogoFinalizado = true;



        removerAviso();

        atualizarLinhaAtual();



        // Pequeno atraso para o jogador ver as cores antes do modal.

        temporizadorModal = setTimeout(

            () => exibirModalFimDeJogo(true),

            400

        );



        return;

    }



    // Próxima tentativa.

    tentativaAtual++;

    letraAtual = 0;



    // Derrota.

    if (tentativaAtual >= linhas.length) {

        jogoFinalizado = true;



        atualizarLinhaAtual();



        // Pequeno atraso para o jogador ver as cores antes do modal.

        temporizadorModal = setTimeout(

            () => exibirModalFimDeJogo(false),

            400

        );



        return;

    }



    atualizarLinhaAtual();

}



// Atualiza a cor das teclas conforme o melhor resultado.

function atualizarCorTecla(letra, novaClasse) {

    const botao = Array.from(botoes).find(

        botao =>

            normalizarPalavra(botao.textContent.trim()) === letra

    );



    if (!botao) {

        return;

    }



    const prioridades = {

        "letra-incorreta": 1,

        "letra-contem": 2,

        "letra-correta": 3

    };



    const classeAtual = botao.dataset.resultado;



    if (

        !classeAtual ||

        prioridades[novaClasse] > prioridades[classeAtual]

    ) {

        botao.dataset.resultado = novaClasse;



        if (novaClasse === "letra-correta") {

            botao.style.backgroundColor = "#13bdb4";

            botao.style.borderColor = "#13bdb4";

        } else if (novaClasse === "letra-contem") {

            botao.style.backgroundColor = "#c7aa2b";

            botao.style.borderColor = "#c7aa2b";

        } else {

            botao.style.backgroundColor = "#777777";

            botao.style.borderColor = "#777777";

        }

    }

}



// Limpa as cores das teclas ao iniciar a partida.

function atualizarTeclado() {

    botoes.forEach(botao => {

        delete botao.dataset.resultado;



        botao.style.backgroundColor = "";

        botao.style.borderColor = "";

    });

}



// Eventos do teclado virtual.

botoes.forEach(botao => {

    if (

        botao.id === "enter" ||

        botao.id === "backspace"

    ) {

        return;

    }



    botao.addEventListener("click", () => {

        inserirLetra(botao.textContent.trim());

    });

});



botaoEnter.addEventListener("click", verificarPalavra);



botaoBackspace.addEventListener("click", apagarLetra);



// Eventos do teclado físico.

document.addEventListener("keydown", evento => {

    if (jogoFinalizado) {

        return;

    }



    if (evento.key === "Enter") {

        evento.preventDefault();

        verificarPalavra();

        return;

    }



    if (evento.key === "Backspace") {

        evento.preventDefault();

        apagarLetra();

        return;

    }



    if (evento.key === "ArrowLeft") {

        evento.preventDefault();

        moverCursor(-1);

        return;

    }



    if (evento.key === "ArrowRight") {

        evento.preventDefault();

        moverCursor(1);

        return;

    }



    if (/^[a-zA-Zá-úÁ-Ú]$/.test(evento.key)) {

        inserirLetra(evento.key);

    }

});



// Inicia somente depois de carregar o dicionário.

carregarDicionario();
