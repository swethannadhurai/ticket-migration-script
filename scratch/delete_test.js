require('dotenv').config();
const axios = require('axios');

const FRESHDESK_DOMAIN = process.env.FRESHDESK_DOMAIN;
const FRESHDESK_API_KEY = process.env.FRESHDESK_API_KEY;
const API_BASE = `https://${FRESHDESK_DOMAIN}/api/v2`;
const AUTH_HEADER = `Basic ${Buffer.from(`${FRESHDESK_API_KEY}:X`).toString('base64')}`;

async function testDelete() {
  const ticketId = 148;
  try {
    // 1. Fetch conversations
    const convos = await axios.get(`${API_BASE}/tickets/${ticketId}/conversations`, {
      headers: { Authorization: AUTH_HEADER }
    });
    
    console.log(`Found ${convos.data.length} conversations on ticket ${ticketId}. IDs: ${convos.data.map(c => c.id).join(', ')}`);
    
    if (convos.data.length === 0) return;
    
    const targetId = convos.data[0].id; // Just try to delete the first one
    console.log(`Trying to delete conversation ${targetId}...`);
    
    // DELETE method
    await axios.delete(`${API_BASE}/conversations/${targetId}`, {
      headers: { Authorization: AUTH_HEADER }
    });
    console.log('✅ Delete via /conversations/{id} SUCCESS!');
    
  } catch (err) {
    if (err.response) {
      console.log(`❌ Delete failed with ${err.response.status}`);
      console.log('Error data:', err.response.data);
    } else {
      console.error(err.message);
    }
  }
}

testDelete();
