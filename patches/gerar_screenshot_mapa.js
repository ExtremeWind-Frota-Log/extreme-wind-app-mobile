/**
 * Gera uma imagem estatica (PNG) do "Mapa de Clientes" (Google My Maps) usando
 * um navegador headless (Puppeteer), e salva em www/assets/mapa-clientes.png.
 *
 * Motivo: o embed do Google My Maps via <iframe> funciona no navegador normal,
 * mas o WebView do app Android (Capacitor) recebe net::ERR_BLOCKED_BY_RESPONSE
 * do Google -- ja tentamos habilitar cookies de terceiros e forcar um
 * User-Agent de Chrome comum, nenhum dos dois resolveu. A solucao estavel e
 * capturar o mapa como imagem (server-side, com um Chrome de verdade) e
 * exibir essa imagem dentro do app, em vez do iframe ao vivo.
 *
 * Este script roda dentro do workflow do GitHub Actions a cada push/build,
 * entao a imagem reflete os marcadores mais recentes do My Maps automaticamente
 * -- nao precisa de atualizacao manual.
 *
 * Uso: node gerar_screenshot_mapa.js
 * (Espera encontrar Puppeteer instalado -- ver step do workflow que faz
 * "npm install puppeteer" antes de chamar este script.)
 */
const path = require("path");
const fs = require("fs");

const MAP_EMBED_URL =
  "https://www.google.com/maps/d/u/0/embed?mid=10Tb-6WWrSx-gk6TIM6r05paJrYBBXcg&ehbc=2E312F";
const OUTPUT_PATH = path.join(__dirname, "..", "www", "assets", "mapa-clientes.png");
const TMP_PATH = OUTPUT_PATH + ".tmp";

// Tamanho do viewport de captura. O card flutuante de Equipes fica sobreposto
// via CSS no index.html (nao faz parte desta imagem), entao a proporcao aqui
// so precisa bater com a altura do bloco do mapa na home (340px) numa largura
// generosa o suficiente pra ficar nitida em telas grandes tambem.
const VIEWPORT = { width: 1280, height: 720 };

// Quantas vezes tentar recarregar a pagina se os tiles nao aparecerem na
// primeira vez (o My Maps as vezes fica preso carregando o fundo cinza).
const MAX_TENTATIVAS = 3;

// Verifica se a screenshot tirada e "util" (nao e so o fundo cinza padrao do
// My Maps enquanto os tiles nao carregam). Conta a proporcao de pixels que
// batem com o cinza de fundo conhecido (~ rgb(229,227,223) claro, ou o tema
// escuro do Maps) -- se quase tudo for a mesma cor solida, provavelmente os
// tiles nao renderizaram.
async function screenshotParecevalido(page) {
  const amostra = await page.evaluate(() => {
    const canvas = document.createElement("canvas");
    canvas.width = 40;
    canvas.height = 40;
    // Nao temos acesso direto aos pixels renderizados via DOM puro; em vez
    // disso, contamos quantos elementos de tile (imagens) o Google Maps ja
    // inseriu no DOM como uma aproximacao barata de "carregou alguma coisa".
    const tiles = document.querySelectorAll('img[src*="maps"], canvas, [class*="tile"]');
    return tiles.length;
  });
  return amostra > 3;
}

async function main() {
  const puppeteer = require("puppeteer");

  const browser = await puppeteer.launch({
    headless: "new",
    args: [
      "--no-sandbox",
      "--disable-setuid-sandbox",
      "--disable-dev-shm-usage",
    ],
  });

  try {
    const page = await browser.newPage();
    await page.setViewport(VIEWPORT);

    let sucesso = false;

    for (let tentativa = 1; tentativa <= MAX_TENTATIVAS && !sucesso; tentativa++) {
      console.log(`Tentativa ${tentativa}/${MAX_TENTATIVAS}: abrindo ${MAP_EMBED_URL} ...`);
      await page.goto(MAP_EMBED_URL, { waitUntil: "networkidle2", timeout: 60000 });

      // O My Maps carrega os tiles de forma assincrona depois do
      // "networkidle2" inicial. Espera fixa generosa para dar tempo dos
      // tiles e marcadores renderizarem visualmente antes do screenshot.
      console.log("Aguardando tiles do mapa renderizarem (10s)...");
      await new Promise((resolve) => setTimeout(resolve, 10000));

      const valido = await screenshotParecevalido(page);
      if (valido) {
        sucesso = true;
        console.log("Tiles do mapa detectados no DOM, seguindo para captura.");
      } else {
        console.log(
          `[aviso] Poucos elementos de tile detectados na tentativa ${tentativa} -- ` +
          "o mapa pode nao ter renderizado visualmente ainda."
        );
      }
    }

    if (!sucesso) {
      throw new Error(
        "Apos " + MAX_TENTATIVAS + " tentativas, o mapa nao parece ter renderizado " +
        "tiles/marcadores visiveis (DOM sem elementos de tile). Abortando sem " +
        "sobrescrever a imagem anterior."
      );
    }

    // Tenta fechar o aviso "Este mapa foi feito com o Google Os Meus Mapas"
    // se ele aparecer, para nao poluir a captura (nao e critico se falhar).
    try {
      const dismissed = await page.evaluate(() => {
        const els = Array.from(document.querySelectorAll("button, a, div"));
        const closeBtn = els.find(
          (el) => el.getAttribute("aria-label") === "Fechar" || el.getAttribute("aria-label") === "Close"
        );
        if (closeBtn) {
          closeBtn.click();
          return true;
        }
        return false;
      });
      if (dismissed) {
        await new Promise((resolve) => setTimeout(resolve, 500));
      }
    } catch (e) {
      console.log("Aviso: nao foi possivel tentar fechar banner do My Maps (nao critico).", e.message);
    }

    // Tira a screenshot num arquivo temporario primeiro. So substitui o
    // arquivo final se a captura realmente tiver variedade de cor (nao e uma
    // tela solida/vazia) -- assim, mesmo se a deteccao de tiles no DOM tiver
    // dado falso positivo, nao sobrescrevemos a imagem boa anterior com uma
    // captura vazia.
    await page.screenshot({ path: TMP_PATH, type: "png" });

    const stats = fs.statSync(TMP_PATH);
    console.log(`Screenshot temporario gerado (${stats.size} bytes). Validando...`);

    if (stats.size < 5000) {
      throw new Error(
        `Screenshot suspeito de estar vazio (apenas ${stats.size} bytes). ` +
        "Abortando sem sobrescrever a imagem anterior."
      );
    }

    fs.renameSync(TMP_PATH, OUTPUT_PATH);
    console.log(`Screenshot do mapa salvo em: ${OUTPUT_PATH}`);
  } finally {
    if (fs.existsSync(TMP_PATH)) {
      fs.unlinkSync(TMP_PATH);
    }
    await browser.close();
  }
}

main().catch((err) => {
  console.error("[ERRO] Falha ao gerar screenshot do mapa:", err);
  // Nao derruba o build inteiro por causa disso -- se falhar, o index.html
  // deve manter a imagem anterior (ou um placeholder) em vez de travar tudo.
  process.exit(1);
});
