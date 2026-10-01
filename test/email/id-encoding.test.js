const os = require('os');
const { callGraphAPI } = require('../../utils/graph-api');
const { ensureAuthenticated } = require('../../auth');
const handleReadEmail = require('../../email/read');
const handleMarkAsRead = require('../../email/mark-as-read');
const handleDeleteEmail = require('../../email/delete');
const { handleDownloadAttachments } = require('../../email/attachments');
const { resolveFolderPath } = require('../../email/folder-utils');
const handleMoveEmails = require('../../folder/move');
const handleDeleteEvent = require('../../calendar/delete');
const handleDeclineEvent = require('../../calendar/decline');
const handleListFiles = require('../../onedrive/list');

jest.mock('../../utils/graph-api', () => ({
  ...jest.requireActual('../../utils/graph-api'),
  callGraphAPI: jest.fn(),
  callGraphAPIPaginated: jest.fn()
}));
jest.mock('../../auth');

// Graph ids are base64 and can contain '=', '+' and '/': each must be encoded exactly once.
const RAW_ID = 'AAMk+ab/cd==';
const ENCODED_ID = 'AAMk%2Bab%2Fcd%3D%3D';

describe('ids in Graph URLs are encoded exactly once', () => {
  beforeEach(() => {
    callGraphAPI.mockReset();
    ensureAuthenticated.mockResolvedValue('dummy_access_token');
    jest.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  const endpoints = () => callGraphAPI.mock.calls.map(call => call[2]);

  test('read-email', async () => {
    callGraphAPI.mockResolvedValueOnce({ id: RAW_ID, subject: 'S', body: { contentType: 'text', content: 'x' } })
      .mockResolvedValueOnce({ value: [] });
    await handleReadEmail({ id: RAW_ID, format: 'html' });
    expect(endpoints()).toEqual([`me/messages/${ENCODED_ID}`, `me/messages/${ENCODED_ID}/attachments`]);
  });

  test('mark-as-read', async () => {
    callGraphAPI.mockResolvedValue({});
    await handleMarkAsRead({ id: RAW_ID });
    expect(endpoints()).toEqual([`me/messages/${ENCODED_ID}`]);
  });

  test('delete-email', async () => {
    callGraphAPI.mockResolvedValue({ id: 'new' });
    await handleDeleteEmail({ id: RAW_ID });
    expect(endpoints()[0]).toBe(`me/messages/${ENCODED_ID}/move`);
  });

  test('download-attachments', async () => {
    callGraphAPI
      .mockResolvedValueOnce({ value: [{ '@odata.type': '#microsoft.graph.fileAttachment', id: RAW_ID, name: 'a.txt' }] })
      .mockResolvedValueOnce({ contentBytes: Buffer.from('a').toString('base64') });
    const fs = require('fs');
    const path = require('path');
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'outlook-mcp-ids-'));
    try {
      await handleDownloadAttachments({ messageId: RAW_ID, saveDir: dir });
    } finally {
      fs.rmSync(dir, { recursive: true, force: true });
    }
    expect(endpoints()).toEqual([
      `me/messages/${ENCODED_ID}/attachments`,
      `me/messages/${ENCODED_ID}/attachments/${ENCODED_ID}`
    ]);
  });

  test('custom folder path for list/search-emails', async () => {
    callGraphAPI.mockResolvedValueOnce({ value: [{ id: RAW_ID, displayName: 'Projekt' }] });
    expect(await resolveFolderPath('dummy_access_token', 'Projekt')).toBe(`me/mailFolders/${ENCODED_ID}/messages`);
  });

  test('folder name with an apostrophe is escaped in the OData filter', async () => {
    callGraphAPI.mockResolvedValueOnce({ value: [{ id: 'f1', displayName: "Ola's" }] });
    await resolveFolderPath('dummy_access_token', "Ola's");
    expect(callGraphAPI.mock.calls[0][4].$filter).toBe("displayName eq 'Ola''s'");
  });

  test('move-emails', async () => {
    callGraphAPI
      .mockResolvedValueOnce({ value: [{ id: 'target', displayName: 'Arkiv' }] }) // folder lookup
      .mockResolvedValue({});
    await handleMoveEmails({ emailIds: RAW_ID, targetFolder: 'Arkiv' });
    expect(endpoints()).toContain(`me/messages/${ENCODED_ID}/move`);
  });

  test('calendar delete and decline', async () => {
    callGraphAPI.mockResolvedValue({});
    await handleDeleteEvent({ eventId: RAW_ID });
    await handleDeclineEvent({ eventId: RAW_ID });
    expect(endpoints()).toEqual([`me/events/${ENCODED_ID}`, `me/events/${ENCODED_ID}/decline`]);
  });

  test('onedrive path segments', async () => {
    callGraphAPI.mockResolvedValue({ value: [] });
    await handleListFiles({ path: '/Mina dokument/a+b/' });
    expect(endpoints()).toEqual(['me/drive/root:/Mina%20dokument/a%2Bb:/children']);
  });
});
