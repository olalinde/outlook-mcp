const fs = require('fs');
const os = require('os');
const path = require('path');
const {
  MAX_ATTACHMENT_BYTES,
  contentTypeFor,
  buildFileAttachments,
  safeFileName
} = require('../../email/attachments');

describe('contentTypeFor', () => {
  test.each([
    ['bild.png', 'image/png'],
    ['bild.PNG', 'image/png'],
    ['foto.jpg', 'image/jpeg'],
    ['foto.jpeg', 'image/jpeg'],
    ['anim.gif', 'image/gif'],
    ['logo.svg', 'image/svg+xml'],
    ['rapport.pdf', 'application/pdf'],
    ['data.xyz', 'application/octet-stream'],
    ['utan-filändelse', 'application/octet-stream']
  ])('%s -> %s', (name, expected) => {
    expect(contentTypeFor(name)).toBe(expected);
  });
});

describe('buildFileAttachments', () => {
  let tmpDir;
  let pngPath;

  beforeAll(() => {
    tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'outlook-mcp-attachments-'));
    pngPath = path.join(tmpDir, 'diagram.png');
    fs.writeFileSync(pngPath, Buffer.from([0x89, 0x50, 0x4e, 0x47, 1, 2, 3]));
  });

  afterAll(() => {
    fs.rmSync(tmpDir, { recursive: true, force: true });
  });

  test('returns an empty array when attachments are missing', () => {
    expect(buildFileAttachments(undefined)).toEqual([]);
    expect(buildFileAttachments(null)).toEqual([]);
  });

  test('builds a fileAttachment with defaults from the file', () => {
    const [attachment] = buildFileAttachments([{ path: pngPath }]);
    expect(attachment).toEqual({
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: 'diagram.png',
      contentType: 'image/png',
      contentBytes: Buffer.from([0x89, 0x50, 0x4e, 0x47, 1, 2, 3]).toString('base64'),
      isInline: false
    });
  });

  test('passes name, contentType, isInline and contentId through', () => {
    const [attachment] = buildFileAttachments([{
      path: pngPath,
      name: 'flode.png',
      contentType: 'image/x-custom',
      isInline: true,
      contentId: 'img1@memento'
    }]);
    expect(attachment.name).toBe('flode.png');
    expect(attachment.contentType).toBe('image/x-custom');
    expect(attachment.isInline).toBe(true);
    expect(attachment.contentId).toBe('img1@memento');
  });

  test('derives contentType from the given name when no contentType is set', () => {
    const [attachment] = buildFileAttachments([{ path: pngPath, name: 'rapport.pdf' }]);
    expect(attachment.contentType).toBe('application/pdf');
  });

  test('rejects files larger than 3 MB', () => {
    const bigPath = path.join(tmpDir, 'big.bin');
    fs.writeFileSync(bigPath, Buffer.alloc(MAX_ATTACHMENT_BYTES + 1));
    expect(() => buildFileAttachments([{ path: bigPath }])).toThrow(/3 MB/);
  });

  test('accepts a file of exactly 3 MB', () => {
    const edgePath = path.join(tmpDir, 'edge.bin');
    fs.writeFileSync(edgePath, Buffer.alloc(MAX_ATTACHMENT_BYTES));
    expect(buildFileAttachments([{ path: edgePath }])).toHaveLength(1);
  });

  test('rejects missing files, relative paths and non-arrays', () => {
    expect(() => buildFileAttachments([{ path: path.join(tmpDir, 'saknas.png') }])).toThrow(/not found/);
    expect(() => buildFileAttachments([{ path: 'relativ.png' }])).toThrow(/absolute/);
    expect(() => buildFileAttachments([{}])).toThrow(/path/);
    expect(() => buildFileAttachments({ path: pngPath })).toThrow(/array/);
  });
});

describe('safeFileName', () => {
  test('replaces characters that are invalid in file names', () => {
    expect(safeFileName('a/b\\c:d*e?f"g<h>i|j.png')).toBe('a_b_c_d_e_f_g_h_i_j.png');
  });

  test('prevents path traversal and empty names', () => {
    expect(safeFileName('../../etc/passwd')).toBe('_.._etc_passwd');
    expect(safeFileName('..')).toBe('attachment');
    expect(safeFileName('')).toBe('attachment');
  });
});
