const { callGraphAPI } = require('../../utils/graph-api');
const { ensureAuthenticated } = require('../../auth');
const handleSearchFiles = require('../../onedrive/search');

jest.mock('../../utils/graph-api');
jest.mock('../../auth');

test("onedrive-search escapes ' in the query as '' and encodes it once", async () => {
  ensureAuthenticated.mockResolvedValue('dummy_access_token');
  callGraphAPI.mockResolvedValue({ value: [] });
  jest.spyOn(console, 'error').mockImplementation(() => {});

  await handleSearchFiles({ query: "Ola's plan" });

  expect(callGraphAPI.mock.calls[0][2]).toBe("me/drive/search(q='Ola''s%20plan')");
  console.error.mockRestore();
});
