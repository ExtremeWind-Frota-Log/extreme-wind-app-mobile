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

const MAP_EMBED_URL =
  "https://www.google.com/maps/d/u/0/embed?mid=10Tb-6WWrSx-gk6TIM6r05paJrYBBXcg&ehbc=2E312F";
const OUTPUT_PATH = path.join(__dirname, "..", "www", "assets", "mapa-clientes.png");

// Tamanho do viewport de captura. O card flutuante de Equipes fica sobreposto
// via CSS no index.html (nao faz parte desta imagem), entao a proporcao aqui
// so precisa bater com a altura do bloco do mapa na home (340px) numa largura
// generosa o suficiente pra ficar nitida em telas grandes tambem.
const VIEWPORT = { width: 1280, height: 720 };

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

    console.log(`Abrindo ${MAP_EMBED_URL} ...`);
    await page.goto(MAP_EMBED_URL, { waitUntil: "networkidle2", timeout: 60000 });

    // O My Maps carrega os tiles de forma assincrona depois do "networkidle2"
    // inicial (o proprio Maps dispara mais requisicoes conforme os tiles
    // entram no viewport). Espera fixa adicional para dar tempo dos tiles e
    // marcadores renderizarem visualmente antes do screenshot.
    console.log("Aguardando tiles do mapa renderizarem...");
    await new Promise((resolve) => setTimeout(resolve, 6000));

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

    await page.screenshot({ path: OUTPUT_PATH, type: "png" });
    console.log(`Screenshot do mapa salvo em: ${OUTPUT_PATH}`);
  } finally {
    await browser.close();
  }
}

main().catch((err) => {
  console.error("[ERRO] Falha ao gerar screenshot do mapa:", err);
  // Nao derruba o build inteiro por causa disso -- se falhar, o index.html
  // deve manter a imagem anterior (ou um placeholder) em vez de travar tudo.
  process.exit(1);
});
