// Universal direct import strategy
const { GoogleGenerativeAI } = require('@google/generative-ai');

module.exports = async (req, res) => {
  // Set explicit CORS headers for serverless execution
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
      return res.status(500).json({ error: 'Config Error: GEMINI_API_KEY is missing from Vercel.' });
    }

    const { prompt } = req.body;
    if (!prompt) {
      return res.status(400).json({ error: 'Request Error: No input prompt data provided.' });
    }

    // Direct instantiation using standard Google SDK guidelines
    const genAI = new GoogleGenerativeAI(apiKey);
    
    // Call the model factory helper directly
    const model = genAI.getGenerativeModel({ model: 'gemini-1.5-flash' });

    const systemInstruction = 
      "You are a master curriculum writer for the Department of Education (DepEd) in the Philippines.\n" +
      "Generate an educational plan based strictly on this JSON format blueprint. Output ONLY valid JSON. Do not wrap in markdown block backticks.\n\n" +
      "Required Format:\n" +
      "{\n" +
      "  \"meta\": { \"topic\": \"\", \"subject\": \"\", \"term\": \"\", \"competencyCode\": \"\", \"duration\": \"60 mins\" },\n" +
      "  \"intentions\": { \"competency\": \"\" },\n" +
      "  \"learningExperiences\": [ { \"step\": 1, \"phase\": \"\", \"duration\": \"\", \"description\": \"\" } ],\n" +
      "  \"slides\": [ { \"slideNum\": 1, \"type\": \"\", \"title\": \"\", \"subtitle\": \"\", \"teacherNote\": \"\" } ]\n" +
      "}";

    const structuredPrompt = `${systemInstruction}\n\nUser Lesson Request:\n${prompt}`;

    // Request text output content generation
    const result = await model.generateContent(structuredPrompt);
    const response = await result.response;
    let text = response.text().trim();

    if (!text) {
      throw new Error('AI engine generated an empty string response.');
    }

    // Strip out accidental markdown block syntax if returned by the model
    if (text.startsWith('```')) {
      text = text.replace(/^```json?/, '').replace(/```$/, '').trim();
    }

    // Validate and parse the clean payload
    let cleanPayload;
    try {
      cleanPayload = JSON.parse(text);
    } catch (parseErr) {
      // Secondary fallback regex check if text is padded
      const jsonExtract = text.match(/\{[\s\S]*\}/);
      if (jsonExtract) {
        cleanPayload = JSON.parse(jsonExtract[0]);
      } else {
        throw new Error('Failed to parse response structure as standard JSON text.');
      }
    }

    return res.status(200).json(cleanPayload);

  } catch (err) {
    console.error('SERVER EXCEPTION DETAILS:', err);
    return res.status(500).json({ 
      error: 'The server processing agent ran into an internal error.',
      details: err.message 
    });
  }
};
