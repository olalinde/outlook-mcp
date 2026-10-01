const { callGraphAPIPaginated } = require('../../utils/graph-api');
const { ensureAuthenticated } = require('../../auth');
const handleSearchEmails = require('../../email/search');

jest.mock('../../utils/graph-api');
jest.mock('../../auth');

const email = (id, subject) => ({
  id, subject, isRead: true, receivedDateTime: '2026-10-01T10:00:00Z',
  from: { emailAddress: { name: 'A', address: 'a@example.com' } }
});

describe('search-emails', () => {
  beforeEach(() => {
    callGraphAPIPaginated.mockReset();
    ensureAuthenticated.mockResolvedValue('dummy_access_token');
    jest.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  const text = (result) => result.content[0].text;

  test('returns matches from the combined search', async () => {
    callGraphAPIPaginated.mockResolvedValueOnce({ value: [email('m1', 'Faktura')] });
    const result = text(await handleSearchEmails({ query: 'faktura' }));
    expect(result).toMatch(/^Found 1 emails matching your search criteria/);
    expect(result).toContain('Faktura');
  });

  test('says nothing matched instead of returning recent emails', async () => {
    callGraphAPIPaginated.mockResolvedValue({ value: [] });
    const result = text(await handleSearchEmails({ query: 'finns-inte', unreadOnly: true }));
    expect(result).toBe('No emails found matching your search criteria.');
    // combined + single-term search only; no "recent emails" or filter-only fallback
    expect(callGraphAPIPaginated).toHaveBeenCalledTimes(2);
    for (const call of callGraphAPIPaginated.mock.calls) {
      expect(call[3].$search).toBeDefined();
    }
  });

  test('labels results that match only part of the criteria', async () => {
    callGraphAPIPaginated
      .mockResolvedValueOnce({ value: [] })                      // combined
      .mockResolvedValueOnce({ value: [] })                      // subject only
      .mockResolvedValueOnce({ value: [email('m2', 'Hej')] });   // from only
    const result = text(await handleSearchEmails({ subject: 'Offert', from: 'a@example.com' }));
    expect(result).toMatch(/^No emails matched all search criteria\. Found 1 emails matching only "from" \(ignored: subject\)/);
  });

  test('reports API errors instead of "no results" when every attempt fails', async () => {
    callGraphAPIPaginated.mockRejectedValue(new Error('API call failed with status 400'));
    const result = text(await handleSearchEmails({ query: 'x' }));
    expect(result).toBe('Error searching emails: API call failed with status 400');
  });
});
