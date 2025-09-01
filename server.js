require('dotenv').config();
const app = require('./src/app');

const PORT = process.env.PORT || 3018;
app.listen(PORT, () => {
  console.log(`[OK] API escuchando en http://localhost:${PORT}`);
});