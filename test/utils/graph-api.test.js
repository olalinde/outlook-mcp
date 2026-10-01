const https = require('https');
const { EventEmitter } = require('events');
const { callGraphAPI, encodePathSegments } = require('../../utils/graph-api');

jest.mock('https');

// Graph ids are base64 and can contain '=', '+' and '/'.
const RAW_ID = 'AAMk+ab/cd==';
const ENCODED_ID = 'AAMk%2Bab%2Fcd%3D%3D';

function mockResponse(statusCode, body) {
  https.request.mockImplementation((url, options, callback) => {
    const res = new EventEmitter();
    res.statusCode = statusCode;
    const req = { on: jest.fn(), write: jest.fn(), end: () => {
      callback(res);
      res.emit('data', JSON.stringify(body));
      res.emit('end');
    } };
    return req;
  });
}

describe('callGraphAPI URL building', () => {
  beforeEach(() => {
    https.request.mockReset();
    jest.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  test('uses an encoded id path as-is, so the id is encoded exactly once', async () => {
    mockResponse(200, { id: RAW_ID });

    await callGraphAPI('token', 'GET', `me/messages/${encodeURIComponent(RAW_ID)}/attachments`, null, { $select: 'id' });

    const url = https.request.mock.calls[0][0];
    expect(url).toBe(`https://graph.microsoft.com/v1.0/me/messages/${ENCODED_ID}/attachments?%24select=id`);
    expect(url).not.toMatch(/%25/); // no double encoding
  });

  test('keeps OneDrive path syntax and encodes each path segment once', async () => {
    mockResponse(200, {});

    await callGraphAPI('token', 'GET', `me/drive/root:/${encodePathSegments('Mina dokument/a+b=c.txt')}:/children`);

    expect(https.request.mock.calls[0][0])
      .toBe('https://graph.microsoft.com/v1.0/me/drive/root:/Mina%20dokument/a%2Bb%3Dc.txt:/children');
  });

  test('uses a full nextLink URL unchanged', async () => {
    mockResponse(200, { value: [] });
    const nextLink = `https://graph.microsoft.com/v1.0/me/mailFolders/${ENCODED_ID}/messages?%24skip=10`;

    await callGraphAPI('token', 'GET', nextLink);

    expect(https.request.mock.calls[0][0]).toBe(nextLink);
  });
});
