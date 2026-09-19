import { expect } from 'chai';
import GoogleSheet from '../src/lib/google-sheet';
import { FakeSheets } from './fake-sheets';

const SPREADSHEET_ID = 'share-spreadsheet';
const SPREADSHEET_TITLE = 'Shared Report';

/**
 * Drive permissions round-trip against the fake: create → list → delete, plus the
 * request shape (sendNotificationEmail, grantee type inference) and validation.
 */
describe('spreadsheet sharing', () => {
  const fake = new FakeSheets();
  let gsheet: GoogleSheet;

  before(() => fake.install());
  after(() => fake.uninstall());

  beforeEach(async () => {
    fake.reset();
    fake.addSpreadsheet(SPREADSHEET_ID, SPREADSHEET_TITLE, [{ title: 'Sheet1' }]);
    gsheet = new GoogleSheet(SPREADSHEET_ID);
    await gsheet.authorize(fake.credentials);
  });

  const driveRequests = (): typeof fake.requests =>
    fake.requests.filter((r) => r.url.includes('/drive/v3/'));

  describe('shareSpreadsheet', () => {
    it('grants a user permission with the requested role', async () => {
      const result = await gsheet.shareSpreadsheet({ emails: ['user@example.com'], role: 'writer' });
      expect(result.granted).to.have.length(1);
      expect(result.granted[0].type).to.equal('user');
      expect(result.granted[0].role).to.equal('writer');
      expect(result.granted[0].emailAddress).to.equal('user@example.com');
      expect(result.granted[0].id).to.match(/^perm-/);
    });

    it('creates one permission per email', async () => {
      const result = await gsheet.shareSpreadsheet({ emails: ['a@x.com', 'b@x.com'], role: 'reader' });
      expect(result.granted).to.have.length(2);
      expect(driveRequests().filter((r) => r.method === 'POST')).to.have.length(2);
    });

    it('defaults notify to false so agents do not spam', async () => {
      await gsheet.shareSpreadsheet({ emails: ['user@example.com'] });
      const request = driveRequests()[0];
      expect(request.url).to.contain('sendNotificationEmail=false');
    });

    it('sends notification email only when asked', async () => {
      await gsheet.shareSpreadsheet({ emails: ['user@example.com'], notify: true, message: 'Report ready' });
      const request = driveRequests()[0];
      expect(request.url).to.contain('sendNotificationEmail=true');
      expect(request.url).to.contain('emailMessage=');
    });

    it('grants a domain permission', async () => {
      const result = await gsheet.shareSpreadsheet({ domain: 'example.com', role: 'commenter' });
      expect(result.granted[0].type).to.equal('domain');
      expect(result.granted[0].domain).to.equal('example.com');
    });

    it('grants an anyone permission', async () => {
      const result = await gsheet.shareSpreadsheet({ anyone: true, role: 'reader' });
      expect(result.granted[0].type).to.equal('anyone');
    });

    it('honors an explicit type override (group)', async () => {
      const result = await gsheet.shareSpreadsheet({ emails: ['group@example.com'], type: 'group' });
      expect(result.granted[0].type).to.equal('group');
    });

    it('rejects a share with no grantee', async () => {
      try {
        await gsheet.shareSpreadsheet({});
        expect.fail('should have thrown');
      } catch (error) {
        expect((error as Error).message).to.contain('emails');
      }
    });

    it('rejects sharing an unknown spreadsheet', async () => {
      try {
        await gsheet.shareSpreadsheet({ emails: ['x@y.com'] }, 'ghost-spreadsheet');
        expect.fail('should have thrown');
      } catch (error) {
        expect((error as Error).message).to.contain('could not find the file');
      }
    });
  });

  describe('listPermissions', () => {
    it('returns the granted permissions', async () => {
      await gsheet.shareSpreadsheet({ emails: ['a@x.com'], role: 'writer' });
      await gsheet.shareSpreadsheet({ emails: ['b@x.com'], role: 'reader' });
      const permissions = await gsheet.listPermissions();
      expect(permissions).to.have.length(2);
      expect(permissions.map((p) => p.emailAddress)).to.eql(['a@x.com', 'b@x.com']);
    });

    it('returns an empty list for an unshared file', async () => {
      const permissions = await gsheet.listPermissions();
      expect(permissions).to.eql([]);
    });
  });

  describe('unshareSpreadsheet', () => {
    it('removes by permissionId', async () => {
      const { granted } = await gsheet.shareSpreadsheet({ emails: ['a@x.com'] });
      const result = await gsheet.unshareSpreadsheet({ permissionId: granted[0].id });
      expect(result.removed).to.equal(true);
      expect(await gsheet.listPermissions()).to.have.length(0);
    });

    it('resolves the permission id from an email', async () => {
      await gsheet.shareSpreadsheet({ emails: ['a@x.com'] });
      const result = await gsheet.unshareSpreadsheet({ email: 'a@x.com' });
      expect(result.removed).to.equal(true);
      expect(await gsheet.listPermissions()).to.have.length(0);
    });

    it('matches email case-insensitively', async () => {
      await gsheet.shareSpreadsheet({ emails: ['User@Example.com'] });
      const result = await gsheet.unshareSpreadsheet({ email: 'user@example.com' });
      expect(result.removed).to.equal(true);
    });

    it('rejects an email with no matching permission', async () => {
      try {
        await gsheet.unshareSpreadsheet({ email: 'nobody@x.com' });
        expect.fail('should have thrown');
      } catch (error) {
        expect((error as Error).message).to.contain('No permission found for "nobody@x.com"');
      }
    });

    it('requires permissionId or email', async () => {
      try {
        await gsheet.unshareSpreadsheet({});
        expect.fail('should have thrown');
      } catch (error) {
        expect((error as Error).message).to.contain('permissionId');
      }
    });
  });

  describe('round-trip', () => {
    it('create → share → list → unshare → empty list', async () => {
      const created = await gsheet.addSpreadsheet('Agent Report');
      const fileId = created.spreadsheetId as string;
      await gsheet.shareSpreadsheet({ emails: ['reviewer@example.com'], role: 'commenter' }, fileId);
      const permissions = await gsheet.listPermissions(fileId);
      expect(permissions).to.have.length(1);
      expect(permissions[0].role).to.equal('commenter');
      await gsheet.unshareSpreadsheet({ permissionId: permissions[0].id }, fileId);
      expect(await gsheet.listPermissions(fileId)).to.have.length(0);
    });
  });
});
