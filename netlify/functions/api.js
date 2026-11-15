const serverless = require('serverless-http');
const { connectLambda } = require('@netlify/blobs');
const { createApp } = require('../../lib/app');
const { blobStorage } = require('../../lib/storage');

const handler = serverless(createApp(blobStorage()));

exports.handler = async (event, context) => {
  connectLambda(event);
  return handler(event, context);
};
