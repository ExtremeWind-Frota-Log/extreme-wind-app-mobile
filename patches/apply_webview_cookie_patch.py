#!/usr/bin/env python3
"""
Patch aplicado no MainActivity.java gerado pelo "npx cap add android" (o
projeto Android e recriado do zero em cada build do GitHub Actions, entao nao
existe MainActivity.java versionado no repo pra editar diretamente -- este
script edita o arquivo depois que o Capacitor ja gerou ele).

Motivo: o iframe do Google My Maps embutido no dashboard de Equipes
("Mapa de Clientes") carrega normalmente no Chrome, mas falha dentro do
WebView do app com "net::ERR_BLOCKED_BY_RESPONSE". Uma causa comum e conhecida
desse erro especifico com Google Maps embed dentro de Android WebView e
cookies de terceiros bloqueados por padrao (CookieManager do Android bloqueia
3rd-party cookies a menos que a app explicite o contrario). Este patch
adiciona a chamada CookieManager.getInstance().setAcceptThirdPartyCookies(...)
apontando pro WebView interno do Capacitor, dentro do onCreate() do
MainActivity.

Isso NAO tem garantia de resolver -- se o bloqueio for do lado do Google
(recusando servir o iframe pra um user-agent de WebView, independente de
cookie), este patch nao muda nada. E uma tentativa de baixo risco: se o
padrao esperado no MainActivity.java nao for encontrado, o script falha com
erro claro (exit code != 0) em vez de aplicar um patch quebrado ou seguir
silenciosamente sem aplicar nada.
"""
import re
import sys

def main():
    if len(sys.argv) != 2:
        print("uso: apply_webview_cookie_patch.py <caminho para MainActivity.java>", file=sys.stderr)
        sys.exit(1)

    path = sys.argv[1]
    with open(path, encoding="utf-8") as f:
        content = f.read()

    if "setAcceptThirdPartyCookies" in content:
        print("Patch ja aplicado anteriormente (idempotente) -- nada a fazer.")
        return

    # MainActivity gerado pelo Capacitor normalmente e so:
    #   package ...;
    #   import com.getcapacitor.BridgeActivity;
    #   public class MainActivity extends BridgeActivity {}
    # (corpo vazio, sem onCreate proprio). Precisamos adicionar um onCreate
    # que chama super.onCreate() e depois habilita os cookies de terceiros
    # no WebView interno (this.bridge.getWebView()).
    class_pattern = re.compile(
        r'(public\s+class\s+MainActivity\s+extends\s+BridgeActivity\s*\{)(\s*)(\})',
        re.S,
    )
    m = class_pattern.search(content)
    if not m:
        print(f"[ERRO] Padrao esperado de MainActivity nao encontrado em {path}. "
              "O patch de cookies de terceiros NAO foi aplicado -- o build "
              "continua, mas o embed do Google Maps pode nao carregar no app. "
              "Revise manualmente o MainActivity.java gerado pelo Capacitor "
              "e ajuste este script se a estrutura da classe mudou.",
              file=sys.stderr)
        sys.exit(1)

    new_body = (
        "\n"
        "  @Override\n"
        "  public void onCreate(android.os.Bundle savedInstanceState) {\n"
        "    super.onCreate(savedInstanceState);\n"
        "    // Habilita cookies de terceiros no WebView do app -- necessario\n"
        "    // pro embed do Google My Maps (dashboard de Equipes) carregar\n"
        "    // dentro do WebView (por padrao o Android bloqueia isso, causando\n"
        "    // net::ERR_BLOCKED_BY_RESPONSE so dentro do app, nao no Chrome).\n"
        "    android.webkit.CookieManager.getInstance().setAcceptThirdPartyCookies(\n"
        "        this.bridge.getWebView(), true);\n"
        "  }\n"
    )
    patched = class_pattern.sub(lambda mm: mm.group(1) + new_body + mm.group(3), content, count=1)

    with open(path, "w", encoding="utf-8") as f:
        f.write(patched)
    print(f"Patch aplicado com sucesso em {path}.")


if __name__ == "__main__":
    main()
