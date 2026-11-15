import serverless from 'serverless-http';
import express from 'express';
import ExcelJS from 'exceljs';
import { getStore } from '@netlify/blobs';
import { createApp } from '../../lib/app.js';
import { blobStorage } from '../../lib/storage.js';

const handler = serverless(createApp(blobStorage()));

// Bridge: Web Request (Netlify Functions v2) -> Lambda-style event -> Express -> Web Response
export default async (req) => {
  const url = new URL(req.url);
  const headers = Object.fromEntries(req.headers);
  const hasBody = !['GET', 'HEAD'].includes(req.method);
  const body = hasBody ? Buffer.from(await req.arrayBuffer()).toString('base64') : null;

  const queryStringParameters = {};
  for (const [k, v] of url.searchParams) queryStringParameters[k] = v;

  const result = await handler(
    {
      httpMethod: req.method,
      path: url.pathname,
      headers,
      queryStringParameters,
      body,
      isBase64Encoded: hasBody,
    },
    {},
  );

  const resBody = result.isBase64Encoded ? Buffer.from(result.body || '', 'base64') : result.body;
  return new Response(resBody, { status: result.statusCode, headers: result.headers });
};

export const config = { path: '/api/*' };
