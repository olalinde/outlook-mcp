const fs = require('fs');
const os = require('os');
const path = require('path');
const { callGraphAPI } = require('../../utils/graph-api');
const { ensureAuthenticated } = require('../../auth');
const { handleDownloadAttachments } = require('../../email/attachments');
const handleSendEmail = require('../../email/send');

jest.mock('../../utils/graph-api');
jest.mock('../../auth');

describe('attachment handlers', () => {
  let tmpDir;

  beforeEach(() => {
    tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'outlook-mcp-handlers-'));
    callGraphAPI.mockReset();
    ensureAuthenticated.mockResolvedValue('dummy_access_token');
    jest.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    fs.rmSync(tmpDir, { recursive: true, force: true });
    console.error.mockRestore();
  });

  test('send-email sends file attachments and leaves the HTML body unchanged', async () => {
    const imgPath = path.join(tmpDir, 'bild.png');
    fs.writeFileSync(imgPath, 'png');
    const body = '<html><body><p style="color:red">Hej</p><img src="cid:img1"></body></html>';
    callGraphAPI.mockResolvedValue({});

    await handleSendEmail({
      to: 'a@example.com',
      subject: 'Test',
      body,
      isHtml: true,
      attachments: [{ path: imgPath, isInline: true, contentId: 'img1' }]
    });

    const [, method, endpoint, payload] = callGraphAPI.mock.calls[0];
    expect(method).toBe('POST');
    expect(endpoint).toBe('me/sendMail');
    expect(payload.message.body).toEqual({ contentType: 'html', content: body });
    expect(payload.message.attachments).toEqual([{
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: 'bild.png',
      contentType: 'image/png',
      contentBytes: Buffer.from('png').toString('base64'),
      isInline: true,
      contentId: 'img1'
    }]);
  });

  test('send-email reports a too large attachment without calling Graph', async () => {
    const bigPath = path.join(tmpDir, 'stor.pdf');
    fs.writeFileSync(bigPath, Buffer.alloc(3 * 1024 * 1024 + 1));

    const result = await handleSendEmail({ to: 'a@example.com', subject: 'S', body: 'B', attachments: [{ path: bigPath }] });

    expect(result.content[0].text).toMatch(/3 MB/);
    expect(callGraphAPI).not.toHaveBeenCalled();
  });

  test('download-attachments saves file attachments and returns metadata', async () => {
    callGraphAPI.mockResolvedValue({
      value: [
        { '@odata.type': '#microsoft.graph.fileAttachment', id: 'a1', name: 'bild.png', contentType: 'image/png', contentId: 'img1', isInline: true, contentBytes: Buffer.from('ett').toString('base64') },
        { '@odata.type': '#microsoft.graph.fileAttachment', id: 'a2', name: 'bild.png', contentType: 'image/png', isInline: true, contentBytes: Buffer.from('två').toString('base64') },
        { '@odata.type': '#microsoft.graph.fileAttachment', id: 'a3', name: 'rapport.pdf', contentType: 'application/pdf', isInline: false, contentBytes: Buffer.from('pdf').toString('base64') },
        { '@odata.type': '#microsoft.graph.itemAttachment', id: 'a4', name: 'Bifogat mejl', isInline: false }
      ]
    });
    const saveDir = path.join(tmpDir, 'ny', 'mapp');

    const result = await handleDownloadAttachments({ messageId: 'msg-1', saveDir, inlineOnly: true });
    const saved = JSON.parse(result.content[0].text);

    expect(callGraphAPI).toHaveBeenCalledWith('dummy_access_token', 'GET', 'me/messages/msg-1/attachments');
    expect(saved).toEqual([
      { path: path.join(saveDir, 'bild.png'), name: 'bild.png', contentType: 'image/png', contentId: 'img1', isInline: true, size: 3 },
      { path: path.join(saveDir, 'bild-1.png'), name: 'bild.png', contentType: 'image/png', contentId: null, isInline: true, size: Buffer.from('två').length }
    ]);
    expect(fs.readFileSync(path.join(saveDir, 'bild.png'), 'utf8')).toBe('ett');
    expect(fs.readFileSync(path.join(saveDir, 'bild-1.png'), 'utf8')).toBe('två');
  });

  test('download-attachments requires an absolute saveDir', async () => {
    const result = await handleDownloadAttachments({ messageId: 'msg-1', saveDir: 'relativ' });
    expect(result.content[0].text).toMatch(/absolute/);
    expect(callGraphAPI).not.toHaveBeenCalled();
  });
});
