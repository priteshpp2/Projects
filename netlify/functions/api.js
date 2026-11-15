const serverless = require('serverless-http');
const { createApp } = require('../../lib/app');
const { blobStorage } = require('../../lib/storage');

const handler = serverless(createApp(blobStorage()));

exports.handler = async (event, context) => handler(event, context);
