// Cloudflare Worker HTMLRewriter AST Pruner
export default {
  async fetch(request, env) {
    const response = await fetch(request);
    const userAgent = request.headers.get("user-agent") || "";
    const isAIBot = /PerplexityBot|GPTBot|ClaudeBot|OAI-SearchBot|Applebot-Extended/i.test(userAgent);
    if (!isAIBot) return response;
    return new HTMLRewriter()
      .on("script:not([type='application/ld+json'])", { element(e) { e.remove(); } })
      .on("svg:not(.critical-icon)", { element(e) { e.remove(); } })
      .on("style, noscript, iframe, canvas", { element(e) { e.remove(); } })
      .on("main, article, [data-chunk-id]", {
        element(e) { e.setAttribute("data-rag-budget", "enforced-14kb"); }
      })
      .transform(response);
  }
};
