// Universal fallback import for Google's Generative AI package
const pkg = require('@google/generative-ai');
const GoogleGenerativeAI = pkg.GoogleGenerativeAI || pkg.GoogleGenAI;

module.exports = async (req, res) => {
  // Setup CORS Headers
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  try {
    const apiKey = process.env.GEMINI_API_KEY;
    if (!apiKey) {
      return res.status(500).json({ error: 'GEMINI_API_KEY is missing from Vercel Environment Variables.' });
    }

    const { prompt } = req.body;
    if (!prompt) {
      return res.status(400).json({ error: 'No prompt found in the request body.' });
    }

    // Initialize using the corrected SDK class
    const ai = new GoogleGenerativeAI(apiKey);
    
    // We target gemini-1.5-flash using the updated method signature
    const model = ai.getGenerativeModel({ 
      model: 'gemini-1.5-flash',
      generationConfig: { responseMimeType: 'application/json' }
    });

    const systemInstruction = 
      "You are a master curriculum writer for the Department of Education (DepEd) in the Philippines. " +
      "Generate an educational framework based strictly on this JSON blueprint. Do not add markdown wrappers or extra text.\n\n" +
      "Structure:\n{\n" +
      "  \"meta\": { \"topic\": \"\", \"subject\": \"\", \"term\": \"\", \"competencyCode\": \"\", \"duration\": \"60 mins\" },\n" +
      "  \"intentions\": { \"competency\": \"\" },\n" +
      "  \"learningExperiences\": [ { \"step\": 1, \"phase\": \"\", \"duration\": \"\", \"description\": \"\" } ],\n" +
      "  \"slides\": [ { \"slideNum\": 1, \"type\": \"\", \"title\": \"\", \"subtitle\": \"\", \"teacherNote\": \"\" } ]\n" +
      "}";

    const combinedPrompt = systemInstruction + "\n\nUser Input Data:\n" + prompt;
    const result = await model.generateContent(combinedPrompt);
    const responseText = result.response && result.response.text ? result.response.text().trim() : '';

    if (!responseText) {
      throw new Error('The AI engine returned a blank response.');
    }

    // Parse output cleanly
    let finalJson;
    try {
      finalJson = JSON.parse(responseText);
    } catch (e) {
      const fallbackMatch = responseText.match(/\{[\s\S]*\}/);
      if (fallbackMatch) {
        finalJson = JSON.parse(fallbackMatch[0]);
      } else {
        throw new Error('AI data structure was not valid JSON.');
      }
    }

    return res.status(200).json(finalJson);

  } catch (err) {
    console.error('SERVER EXCEPTION:', err);
    return res.status(500).json({ 
      error: 'The server processing agent ran into an internal error.',
      details: err.message 
    });
  }
};
