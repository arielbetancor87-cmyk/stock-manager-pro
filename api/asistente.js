// Función serverless de Vercel — intermediario seguro entre la app y la
// API de Anthropic. La API key vive solo acá (variable de entorno del
// servidor), nunca llega al navegador del usuario.
export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Método no permitido" });
  }

  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) {
    return res.status(500).json({ error: "Falta configurar ANTHROPIC_API_KEY en Vercel" });
  }

  const { pregunta, contexto } = req.body || {};
  if (!pregunta || typeof pregunta !== "string") {
    return res.status(400).json({ error: "Falta la pregunta" });
  }

  const systemPrompt = `Sos el asistente de datos de StockPro (Venta Directa), una app de venta por catálogo.
Respondé SOLO en base al contexto JSON que se te da a continuación — nunca inventes números.
Si el contexto no alcanza para responder, decilo con claridad.
Respondé en español, corto y directo, en 2-4 oraciones. Usá $ con separador de miles para montos en pesos argentinos.
Si corresponde, cerrá con una lista breve (máximo 5 ítems) usando guiones.

CONTEXTO ACTUAL DEL NEGOCIO:
${JSON.stringify(contexto || {}, null, 2)}`;

  try {
    const r = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": apiKey,
        "anthropic-version": "2023-06-01",
      },
      body: JSON.stringify({
        model: "claude-sonnet-4-6",
        max_tokens: 500,
        system: systemPrompt,
        messages: [{ role: "user", content: pregunta }],
      }),
    });

    if (!r.ok) {
      const errText = await r.text();
      console.error("Error de Anthropic:", errText);
      return res.status(502).json({ error: "Error consultando la IA" });
    }

    const data = await r.json();
    const texto = (data.content || []).map((b) => b.text || "").join("\n").trim();
    return res.status(200).json({ respuesta: texto || "No pude generar una respuesta." });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: "Error interno" });
  }
}
