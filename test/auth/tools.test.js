const { tokenStorage, handleCheckAuthStatus, handleAuthenticate } = require('../../auth/tools');

jest.mock('../../auth/token-storage'); // automock: tokenStorage methods are jest.fn()

describe('check-auth-status', () => {
  beforeEach(() => {
    jest.resetAllMocks();
    jest.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  const text = (result) => result.content[0].text;

  test('reports not authenticated when there are no tokens', async () => {
    tokenStorage.getTokens.mockResolvedValue(null);
    expect(text(await handleCheckAuthStatus())).toMatch(/^Not authenticated/);
    expect(tokenStorage.getValidAccessToken).not.toHaveBeenCalled();
  });

  test('reports authenticated with expiry when the access token is valid', async () => {
    const expiresAt = Date.now() + 3600000;
    tokenStorage.getTokens.mockResolvedValue({ access_token: 'a', refresh_token: 'r', expires_at: expiresAt });
    tokenStorage.isTokenExpired.mockReturnValue(false);
    tokenStorage.getValidAccessToken.mockResolvedValue('a');
    tokenStorage.getExpiryTime.mockReturnValue(expiresAt);

    const result = text(await handleCheckAuthStatus());
    expect(result).toMatch(/^Authenticated and ready\./);
    expect(result).toContain(new Date(expiresAt).toISOString());
  });

  test('refreshes an expired access token instead of reporting not authenticated', async () => {
    const newExpiry = Date.now() + 3600000;
    tokenStorage.getTokens.mockResolvedValue({ access_token: 'old_token_value', refresh_token: 'r', expires_at: Date.now() - 1000 });
    tokenStorage.isTokenExpired.mockReturnValue(true);
    tokenStorage.getValidAccessToken.mockResolvedValue('new_token_value');
    tokenStorage.getExpiryTime.mockReturnValue(newExpiry);

    const result = text(await handleCheckAuthStatus());
    expect(tokenStorage.getValidAccessToken).toHaveBeenCalled();
    expect(result).toMatch(/^Authenticated and ready \(access token was expired and has been refreshed\)/);
    expect(result).not.toMatch(/token_value/); // token values are never shown
  });

  test('reports a failed refresh', async () => {
    tokenStorage.getTokens.mockResolvedValue({ access_token: 'old', refresh_token: 'r', expires_at: Date.now() - 1000 });
    tokenStorage.isTokenExpired.mockReturnValue(true);
    tokenStorage.getValidAccessToken.mockResolvedValue(null);

    expect(text(await handleCheckAuthStatus())).toMatch(/^Not authenticated: .*refresh failed/);
  });
});

describe('authenticate', () => {
  beforeEach(() => {
    jest.resetAllMocks();
    jest.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  test('starts the device code flow and returns the code and URL', async () => {
    tokenStorage.startDeviceCode.mockResolvedValue({
      user_code: 'ABCD-1234', device_code: 'dev', verification_uri: 'https://microsoft.com/devicelogin', interval: 5, expires_in: 900
    });
    tokenStorage.pollDeviceCode.mockResolvedValue({});

    const result = await handleAuthenticate({});

    expect(result.content[0].text).toContain('https://microsoft.com/devicelogin');
    expect(result.content[0].text).toContain('ABCD-1234');
    expect(tokenStorage.pollDeviceCode).toHaveBeenCalledWith('dev', 5, 900);
    expect(tokenStorage.clearTokens).not.toHaveBeenCalled();
  });

  test('force clears the stored tokens first', async () => {
    tokenStorage.startDeviceCode.mockResolvedValue({ user_code: 'X', device_code: 'd', verification_uri: 'u' });
    tokenStorage.pollDeviceCode.mockResolvedValue({});

    await handleAuthenticate({ force: true });

    expect(tokenStorage.clearTokens).toHaveBeenCalled();
  });
});
