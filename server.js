const path = require('path');
const express = require('express');
const { createApp } = require('./lib/app');
const { pickStorage } = require('./lib/storage');

const PORT = process.env.PORT || 3000;
const storage = pickStorage();

const app = createApp(storage);
app.use(express.static(path.join(__dirname, 'public')));

app.listen(PORT, () => {
  console.log(`\n  Streak App running  ->  http://localhost:${PORT}`);
  console.log(`  Data storage        ->  ${storage.describe()}\n`);
});
