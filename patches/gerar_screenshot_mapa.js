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

// Numero minimo de requisicoes de imagem de tile (gstatic/maps) que precisam
// ter respondido com sucesso antes de considerarmos o mapa "carregado". Um
// mapa vazio/cinza nao dispara essas requisicoes; um mapa real dispara
// dezenas delas conforme os tiles do viewport chegam.
const MIN_TILES_CARREGADOS = 15;

// Verifica se a pagina parece ter renderizado algo visual (nao e so o fundo
// cinza padrao do My Maps enquanto os tiles nao carregam). Combina duas
// evidencias: (1) contagem de requisicoes de rede de tiles de imagem que
// completaram com sucesso (evidencia direta de que o Maps buscou o
// conteudo), e (2) tamanho do PNG resultante como reforço (tiles reais
// aumentam bastante o tamanho do arquivo comparado a uma tela solida).
async function screenshotParecevalido(page, caminhoTemp, tilesCarregados) {
  await page.screenshot({ path: caminhoTemp, type: "png" });
  const stats = fs.statSync(caminhoTemp);
  console.log(
    `  (verificacao: ${tilesCarregados} tiles carregados via rede, ` +
    `screenshot temporaria tem ${stats.size} bytes)`
  );
  return tilesCarregados >= MIN_TILES_CARREGADOS && stats.size > 15000;
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

    // Diagnostico extra: loga erros de console/pagina e falhas de request
    // (ex: bloqueio por CORS, CSP, ou recusa do Google -- o mesmo tipo de
    // problema que afeta o WebView do Android pode, em tese, tambem afetar
    // um Chrome headless "de servidor" se o Google decidir bloquear por
    // outro sinal, como IP de datacenter/user-agent headless).
    page.on("console", (msg) => {
      const tipo = msg.type();
      if (tipo === "error" || tipo === "warning") {
        console.log(`    [console.${tipo}] ${msg.text().slice(0, 300)}`);
      }
    });
    page.on("requestfailed", (req) => {
      console.log(`    [request-falhou] ${req.resourceType()} ${req.url().slice(0, 150)} -- ${req.failure()?.errorText}`);
    });
    page.on("pageerror", (err) => {
      console.log(`    [pageerror] ${String(err).slice(0, 300)}`);
    });

    // Conta requisicoes de rede de tiles de imagem do Google Maps que
    // respondem com sucesso (status 200-299). Isso e uma evidencia direta de
    // que o mapa esta buscando/recebendo os tiles visuais, independente de
    // qual elemento de DOM o Google usa para renderiza-los.
    //
    // Diagnostico (build #167): o filtro anterior (khms / maps.gstatic.com /
    // vt) deu 0 em 3 tentativas seguidas, mesmo com uma captura de 19497
    // bytes (mais que o placeholder de 11014). Em vez de adivinhar outro
    // hostname, agora logamos TODA resposta de imagem/rede relevante (sem
    // filtro de host) para ver exatamente o que o embed carrega, e usamos um
    // filtro bem mais amplo (qualquer resposta de imagem de dominio google)
    // como evidencia de tile.
    let tilesCarregados = 0;
    const urlsVistas = new Set();
    page.on("response", (response) => {
      try {
        const url = response.url();
        const req = response.request();
        const tipo = req.resourceType();
        const isGoogleImg =
          tipo === "image" &&
          /google|gstatic|ggpht|googleusercontent/i.test(url);
        if (isGoogleImg && response.ok()) {
          tilesCarregados++;
        }
        // Log de diagnostico: registra dominios distintos vistos (nao a URL
        // inteira, pra nao poluir o log com querystrings gigantes), so a
        // primeira vez que cada host aparece.
        try {
          const host = new URL(url).host;
          const chave = `${tipo}:${host}`;
          if (!urlsVistas.has(chave)) {
            urlsVistas.add(chave);
            console.log(`    [rede] tipo=${tipo} host=${host} status=${response.status()}`);
          }
        } catch (e) {
          // URL invalida, ignora o log de diagnostico.
        }
      } catch (e) {
        // Ignora erros de leitura de resposta (conexao fechada etc).
      }
    });

    let sucesso = false;
    // Guarda a melhor tentativa (maior numero de tiles carregados) mesmo se
    // nenhuma bater o limiar ideal -- assim, se o limiar estiver calibrado
    // errado, ainda usamos a captura mais completa em vez de descartar tudo
    // e ficar preso no placeholder para sempre.
    let melhorTiles = -1;
    let melhorBytes = 0;
    const MELHOR_TMP_PATH = OUTPUT_PATH + ".melhor.tmp";

    for (let tentativa = 1; tentativa <= MAX_TENTATIVAS && !sucesso; tentativa++) {
      tilesCarregados = 0;
      console.log(`Tentativa ${tentativa}/${MAX_TENTATIVAS}: abrindo ${MAP_EMBED_URL} ...`);
      await page.goto(MAP_EMBED_URL, { waitUntil: "networkidle2", timeout: 60000 });

      // O My Maps carrega os tiles de forma assincrona depois do
      // "networkidle2" inicial. Espera fixa generosa para dar tempo dos
      // tiles e marcadores renderizarem visualmente antes do screenshot.
      console.log("Aguardando tiles do mapa renderizarem (15s)...");
      await new Promise((resolve) => setTimeout(resolve, 15000));

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

      // Diagnostico extra: titulo da pagina, texto visivel (pode conter uma
      // mensagem de erro do Google, como "nao e possivel carregar" etc), e
      // contagem de elementos <img>/<canvas> no DOM (independente de rede).
      try {
        const diag = await page.evaluate(() => ({
          titulo: document.title,
          textoInicial: (document.body.innerText || "").slice(0, 200).replace(/\s+/g, " "),
          numImgs: document.querySelectorAll("img").length,
          numCanvas: document.querySelectorAll("canvas").length,
          numIframes: document.querySelectorAll("iframe").length,
        }));
        console.log(
          `    [diag-dom] titulo="${diag.titulo}" imgs=${diag.numImgs} canvas=${diag.numCanvas} ` +
          `iframes=${diag.numIframes} texto="${diag.textoInicial}"`
        );
      } catch (e) {
        console.log("    [diag-dom] falhou ao ler DOM:", e.message);
      }

      const valido = await screenshotParecevalido(page, TMP_PATH, tilesCarregados);
      const stats = fs.statSync(TMP_PATH);

      if (tilesCarregados > melhorTiles) {
        melhorTiles = tilesCarregados;
        melhorBytes = stats.size;
        fs.copyFileSync(TMP_PATH, MELHOR_TMP_PATH);
      }

      if (valido) {
        sucesso = true;
        console.log("Screenshot com tiles suficientes carregados, aceitando.");
      } else {
        console.log(
          `[aviso] Screenshot da tentativa ${tentativa} parece incompleta -- ` +
          "o mapa pode nao ter renderizado visualmente ainda."
        );
      }
    }

    if (!sucesso) {
      // Piso minimo de seguranca: exige que a melhor tentativa seja
      // claramente maior que o placeholder conhecido (~11KB), mesmo sem
      // bater o limiar ideal de tiles -- evita publicar algo pior que o que
      // ja existe, mas tambem evita ficar travado no placeholder para
      // sempre se o limiar ideal estiver calibrado alto demais.
      if (melhorTiles > 0 && melhorBytes > 13000 && fs.existsSync(MELHOR_TMP_PATH)) {
        console.log(
          `[aviso] Nenhuma tentativa bateu o limiar ideal (${MIN_TILES_CARREGADOS} tiles), ` +
          `mas a melhor tentativa teve ${melhorTiles} tiles / ${melhorBytes} bytes -- usando-a ` +
          "como melhor esforco, em vez de manter o placeholder."
        );
        fs.copyFileSync(MELHOR_TMP_PATH, OUTPUT_PATH);
        fs.unlinkSync(MELHOR_TMP_PATH);
        console.log(`Screenshot do mapa (melhor esforco) salvo em: ${OUTPUT_PATH}`);
        return;
      }
      if (fs.existsSync(MELHOR_TMP_PATH)) {
        fs.unlinkSync(MELHOR_TMP_PATH);
      }
      throw new Error(
        `Apos ${MAX_TENTATIVAS} tentativas, a melhor captura teve apenas ${melhorTiles} ` +
        `tiles / ${melhorBytes} bytes -- proximo demais do placeholder vazio. ` +
        "Abortando sem sobrescrever a imagem anterior."
      );
    }

    if (fs.existsSync(MELHOR_TMP_PATH)) {
      fs.unlinkSync(MELHOR_TMP_PATH);
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
