const { test, expect } = require('@playwright/test');

test.describe('Eco-growth Discovery Taskpane UI Tests', () => {
  test.beforeEach(async ({ page }) => {
    // 1. Intercept network request to Microsoft's hosted office.js script
    // Fulfill with a comment-only string so the global Office object
    // injected via addInitScript is used.
    await page.route(
      'https://appsforoffice.microsoft.com/lib/1/hosted/office.js',
      async (route) => {
        await route.fulfill({
          contentType: 'application/javascript',
          body: '// Mock Office.js library',
        });
      }
    );

    // 2. Mock Office.js globals
    await page.addInitScript(() => {
      window.Office = {
        HostType: { PowerPoint: 'PowerPoint' },
        FileType: { Compressed: 'Compressed' },
        AsyncResultStatus: { Succeeded: 'Succeeded', Failed: 'Failed' },

        onReady: (callback) => {
          // Delay callback to allow taskpane page to initialize
          setTimeout(() => {
            callback({ host: 'PowerPoint' });
          }, 100);
        },

        context: {
          document: {
            getFileAsync: (fileType, options, callback) => {
              // Mock getFileAsync behavior with 2 slices, delaying to capture "Reading active file(s)..."
              setTimeout(() => {
                callback({
                  status: 'Succeeded',
                  value: {
                    size: 131072, // 128 KB
                    sliceCount: 2,
                    getSliceAsync: (sliceIndex, sliceCallback) => {
                      // Delay to allow Playwright to capture intermediate state (progress)
                      setTimeout(() => {
                        sliceCallback({
                          status: 'Succeeded',
                          value: {
                            data: new Uint8Array(65536), // 64 KB slice data
                          },
                        });
                      }, 500);
                    },
                    closeAsync: (closeCallback) => {
                      closeCallback();
                    },
                  },
                });
              }, 500);
            },
          },
        },
      };
    });

    // 3. Navigate to taskpane
    await page.goto('/');
  });

  test('initializes and loads JV initials', async ({ page }) => {
    // The element defaults to '--' and changes to 'JV' on Office initialization
    const initialsDisplay = page.locator('#initials-display');
    await expect(initialsDisplay).toHaveText('JV');
  });

  test('reads active files and updates status progress and success', async ({ page }) => {
    // Verify initials JV are loaded first to ensure Office.js onReady has completed
    const initialsDisplay = page.locator('#initials-display');
    await expect(initialsDisplay).toHaveText('JV');

    const readBtn = page.locator('#read-files-btn');
    await expect(readBtn).toHaveText('Read Active File(s)');

    const status = page.locator('#status');
    await expect(status).toBeEmpty();

    // Trigger read action
    await readBtn.click();

    // The status should transition to 'Reading active file(s)...'
    await expect(status).toHaveText('Reading active file(s)...');

    // Wait for file size and slice details
    await expect(status).toHaveText(/File size: 131072 bytes. Reading 2 slices.../);

    // Then progress states (50% or 100%) and then final success message
    await expect(status).toHaveText(/(Reading progress: (50|100)%|Successfully read active file\(s\): 131072 bytes.)/);

    // Finally verify complete success
    await expect(status).toHaveText('Successfully read active file(s): 131072 bytes.');
  });
});
