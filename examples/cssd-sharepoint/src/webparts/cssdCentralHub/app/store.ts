/* eslint-disable @typescript-eslint/no-explicit-any */
// SharePoint persistence for CSSD Central Hub.
//
// Every collection lives in its own SharePoint list so site owners can set
// permissions per list (for example: management lists for the In-Charge group only).
// Each list item holds the full record as JSON in CssdRecord, plus a few
// promoted columns so the lists stay readable in SharePoint, Excel and Power BI.
// Attachments are uploaded to the "CSSD Attachments" document library.

export interface HttpResponseLike {
  ok: boolean;
  status: number;
  // eslint-disable-next-line @rushstack/no-new-null
  headers: { get(name: string): string | null };
  json(): Promise<any>;
  text(): Promise<string>;
}

export interface HttpLike {
  get(url: string): Promise<HttpResponseLike>;
  post(url: string, body?: string | ArrayBuffer, headers?: Record<string, string>): Promise<HttpResponseLike>;
}

type ColumnType = 'Text' | 'Boolean';

interface ColumnDef {
  name: string;
  display: string;
  type: ColumnType;
  field: string;
}

export interface ListDef {
  key: string;
  url: string;
  title: string;
  description: string;
  columns: ColumnDef[];
}

const text = (name: string, display: string, field: string): ColumnDef => ({ name, display, type: 'Text', field });

export const LISTS: ListDef[] = [
  { key: 'orders', url: 'CSSDOrders', title: 'CSSD Orders', description: 'Clinic and OR orders sent to CSSD and their receiving records.',
    columns: [text('CssdDepartment', 'Department', 'dept'), text('CssdOrderType', 'Order type', 'type'), text('CssdDate', 'Date', 'date'),
      text('CssdShift', 'Shift', 'shift'), text('CssdStatus', 'Status', 'status'), { name: 'CssdRush', display: 'Rush', type: 'Boolean', field: 'rush' },
      text('CssdSubmittedBy', 'Submitted by', 'by'), text('CssdReceiver', 'CSSD receiver', 'receiver')] },
  { key: 'issues', url: 'CSSDIssues', title: 'CSSD Issues', description: 'Department reports and issue follow-up.',
    columns: [text('CssdIssueTitle', 'Issue title', 'title'), text('CssdIssueType', 'Issue type', 'type'), text('CssdPriority', 'Priority', 'priority'),
      text('CssdStatus', 'Status', 'status'), text('CssdReportedBy', 'Reported by', 'author'), text('CssdDate', 'Date', 'date')] },
  { key: 'loaners', url: 'CSSDLoanerSets', title: 'CSSD Loaner Sets', description: 'Loaner set receiving, processing, delivery and return.',
    columns: [text('CssdCompany', 'Company', 'company'), text('CssdSetName', 'Set', 'set'), text('CssdStatus', 'Status', 'status'),
      text('CssdReceivedAt', 'Received at', 'date'), text('CssdReceiver', 'Receiver', 'receiver')] },
  { key: 'endorsements', url: 'CSSDShiftHandovers', title: 'CSSD Shift Handovers', description: 'In-charge shift handover records.',
    columns: [text('CssdInCharge', 'In-charge', 'name'), text('CssdDate', 'Date', 'date'), text('CssdShift', 'Shift', 'shift'),
      text('CssdNextShift', 'Receiving staff (next shift)', 'receiver')] },
  { key: 'parameters', url: 'CSSDParameters', title: 'CSSD Parameters', description: 'Device parameter records entered by staff.',
    columns: [text('CssdStaff', 'Staff', 'name'), text('CssdDevice', 'Device / Area', 'device'), text('CssdParameter', 'Parameter', 'parameter'),
      text('CssdValue', 'Value', 'value'), text('CssdResult', 'Result', 'result'), text('CssdDate', 'Date', 'date')] },
  { key: 'tasks', url: 'CSSDStaffTasks', title: 'CSSD Staff Tasks', description: 'Staff task assignments from department management.',
    columns: [text('CssdTask', 'Task', 'title'), text('CssdAssignedTo', 'Assigned to', 'owner'), text('CssdDate', 'Date', 'date'),
      text('CssdShift', 'Shift', 'shift'), text('CssdStatus', 'Status', 'status')] },
  { key: 'devices', url: 'CSSDDeviceStatus', title: 'CSSD Device Status', description: 'Manually recorded device readiness.',
    columns: [text('CssdDevice', 'Device', 'name'), text('CssdStatus', 'Status', 'status'), text('CssdUpdatedBy', 'Updated by', 'author')] },
  { key: 'ifuDocuments', url: 'CSSDIFULibrary', title: 'CSSD IFU Library', description: 'Instructions for use (IFU) per instrument.',
    columns: [text('CssdInstrument', 'Instrument', 'title'), text('CssdManufacturer', 'Company', 'manufacturer'), text('CssdModel', 'Model or REF', 'model')] },
  { key: 'emergencies', url: 'CSSDEmergencyCases', title: 'CSSD Emergency Cases', description: 'Emergency instrument and set requests and their follow-up.',
    columns: [text('CssdDepartment', 'Department', 'dept'), text('CssdSetName', 'Instrument or set', 'set'), text('CssdStatus', 'Status', 'status'),
      text('CssdRequestTime', 'Request time', 'at')] },
  { key: 'recalls', url: 'CSSDRecalls', title: 'CSSD Recalls', description: 'Recall reports, decisions and OR follow-up.',
    columns: [text('CssdReason', 'Reason', 'reason'), text('CssdLoad', 'Load', 'load'), text('CssdDevice', 'Device', 'device'),
      text('CssdProcessDate', 'Processing date', 'processDate'), text('CssdStatus', 'Status', 'status')] }
];

export const ATTACHMENTS = { url: 'CSSDAttachments', title: 'CSSD Attachments', description: 'Files attached to CSSD Central Hub records.' };

// SharePoint plain multi-line text columns are reliable up to this length.
const MAX_RECORD_CHARS = 60000;

export interface HubState {
  orders: any[];
  issues: any[];
  loaners: any[];
  endorsements: any[];
  parameters: any[];
  ifuDocuments: any[];
  emergencies: any[];
  recalls: any[];
  management: { tasks: any[]; devices: any[] };
}

function rowsOf(state: HubState, key: string): any[] {
  if (key === 'tasks' || key === 'devices') return state.management[key];
  return (state as any)[key];
}

interface Snapshot { itemId: number; etag: string; json: string }

export class ConflictError extends Error {}

const odataString = (value: string): string => encodeURIComponent(`'${value.replace(/'/g, "''")}'`);

export class HubStore {
  private snaps = new Map<string, Map<string, Snapshot>>();
  private chain: Promise<void> = Promise.resolve();

  public constructor(private http: HttpLike, private webUrl: string, private webPath: string) {
    this.webPath = webPath.replace(/\/$/, '');
  }

  private listPath(def: { url: string }, library = false): string {
    return `${this.webPath}${library ? '' : '/Lists'}/${def.url}`;
  }

  private listApi(def: ListDef, suffix: string, query = ''): string {
    return `${this.webUrl}/_api/web/GetList(@a)${suffix}?@a=${odataString(this.listPath(def))}${query}`;
  }

  private async fail(res: HttpResponseLike, what: string): Promise<never> {
    let detail = '';
    try {
      const body = await res.json();
      detail = body?.error?.message?.value || body?.['odata.error']?.message?.value || body?.error?.message || '';
    } catch { /* body was not JSON */ }
    if (res.status === 412) throw new ConflictError('Another user changed this record. The latest records were reloaded.');
    if (res.status === 401 || res.status === 403) throw new Error(`${what}: you do not have permission on this SharePoint site.`);
    throw new Error(`${what} failed (${res.status})${detail ? ': ' + detail : ''}`);
  }

  /** Lists (and the attachments library) that do not exist yet. */
  public async missingLists(): Promise<string[]> {
    const checks = [...LISTS.map(d => ({ title: d.title, path: this.listPath(d) })), { title: ATTACHMENTS.title, path: this.listPath(ATTACHMENTS, true) }];
    const missing: string[] = [];
    for (const c of checks) {
      const res = await this.http.get(`${this.webUrl}/_api/web/GetList(@a)?@a=${odataString(c.path)}&$select=Id`);
      if (res.status === 404) missing.push(c.title);
      else if (!res.ok) await this.fail(res, `Checking ${c.title}`);
    }
    return missing;
  }

  /** Creates the missing lists, columns and the attachments library. Requires Manage Lists permission. */
  public async provision(progress: (message: string) => void): Promise<void> {
    const missing = new Set(await this.missingLists());
    for (const def of LISTS) {
      if (!missing.has(def.title)) continue;
      progress(`Creating ${def.title}…`);
      const id = await this.createList(def.url, def.title, def.description, 100);
      const base = `${this.webUrl}/_api/web/lists(guid'${id}')`;
      await this.addField(base, `<Field Type="Note" Name="CssdRecord" StaticName="CssdRecord" DisplayName="CSSD record (JSON)" NumLines="6" RichText="FALSE" UnlimitedLengthInDocumentLibrary="TRUE" />`, false);
      for (const c of def.columns) {
        const xml = c.type === 'Boolean'
          ? `<Field Type="Boolean" Name="${c.name}" StaticName="${c.name}" DisplayName="${c.display}"><Default>0</Default></Field>`
          : `<Field Type="Text" Name="${c.name}" StaticName="${c.name}" DisplayName="${c.display}" MaxLength="255" />`;
        await this.addField(base, xml, true);
      }
      const res = await this.http.post(`${base}/fields/getbyinternalnameortitle('Title')`, JSON.stringify({ Title: 'Record ID', Indexed: true, EnforceUniqueValues: true }), { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' });
      if (!res.ok) await this.fail(res, `Configuring ${def.title}`);
    }
    if (missing.has(ATTACHMENTS.title)) {
      progress(`Creating ${ATTACHMENTS.title}…`);
      await this.createList(ATTACHMENTS.url, ATTACHMENTS.title, ATTACHMENTS.description, 101);
    }
  }

  private async createList(url: string, title: string, description: string, template: number): Promise<string> {
    // Create with the short name first so the URL stays clean, then set the display title.
    const res = await this.http.post(`${this.webUrl}/_api/web/lists`, JSON.stringify({ BaseTemplate: template, Title: url, Description: description, ContentTypesEnabled: false }));
    if (!res.ok) await this.fail(res, `Creating ${title}`);
    const id: string = (await res.json()).Id;
    const rename = await this.http.post(`${this.webUrl}/_api/web/lists(guid'${id}')`, JSON.stringify({ Title: title }), { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' });
    if (!rename.ok) await this.fail(rename, `Renaming ${title}`);
    return id;
  }

  private async addField(listBase: string, schemaXml: string, addToDefaultView: boolean): Promise<void> {
    const options = 8 + (addToDefaultView ? 16 : 0); // AddFieldInternalNameHint (+ AddToDefaultView)
    const res = await this.http.post(`${listBase}/fields/createfieldasxml`, JSON.stringify({ parameters: { SchemaXml: schemaXml, Options: options } }));
    if (!res.ok) await this.fail(res, 'Adding a column');
  }

  /** Reads every list into a fresh state object and resets the change snapshots. */
  public async load(): Promise<HubState> {
    await this.idle();
    const state: HubState = { orders: [], issues: [], loaners: [], endorsements: [], parameters: [], ifuDocuments: [], emergencies: [], recalls: [], management: { tasks: [], devices: [] } };
    const snaps = new Map<string, Map<string, Snapshot>>();
    await Promise.all(LISTS.map(async def => {
      const snap = new Map<string, Snapshot>();
      const rows = rowsOf(state, def.key);
      let url: string | undefined = this.listApi(def, '/items', '&$select=Id,Title,CssdRecord&$orderby=Id desc&$top=2000');
      while (url) {
        const res = await this.http.get(url);
        if (!res.ok) await this.fail(res, `Reading ${def.title}`);
        const body = await res.json();
        for (const item of body.value || []) {
          let record: any;
          try { record = JSON.parse(item.CssdRecord); } catch { continue; }
          if (!record || typeof record !== 'object' || record.id === undefined) continue;
          rows.push(record);
          snap.set(String(record.id), { itemId: item.Id, etag: item['@odata.etag'] || item['odata.etag'] || '*', json: item.CssdRecord });
        }
        url = body['@odata.nextLink'] || body['odata.nextLink'];
      }
      // Orders are shown oldest first; every other collection newest first.
      if (def.key === 'orders') rows.reverse();
      snaps.set(def.key, snap);
    }));
    this.snaps = snaps;
    return state;
  }

  /** Writes what changed since the last load or sync. Calls are serialized. */
  public sync(state: HubState): Promise<void> {
    const run = this.chain.then(() => this.writeChanges(state));
    this.chain = run.catch(() => undefined);
    return run;
  }

  public idle(): Promise<void> {
    return this.chain;
  }

  private async writeChanges(state: HubState): Promise<void> {
    for (const def of LISTS) {
      const snap = this.snaps.get(def.key) || new Map<string, Snapshot>();
      this.snaps.set(def.key, snap);
      const seen = new Set<string>();
      for (const record of rowsOf(state, def.key)) {
        const key = String(record.id);
        seen.add(key);
        const json = JSON.stringify(record);
        const before = snap.get(key);
        if (before && before.json === json) continue;
        if (json.length > MAX_RECORD_CHARS) throw new Error('This record is too large to save. Shorten the notes or split it into separate records.');
        const body: Record<string, unknown> = { Title: key, CssdRecord: json };
        for (const c of def.columns) {
          const v = record[c.field];
          body[c.name] = c.type === 'Boolean' ? Boolean(v) : String(v ?? '').slice(0, 255);
        }
        if (!before) {
          const res = await this.http.post(this.listApi(def, '/items'), JSON.stringify(body));
          if (!res.ok) await this.fail(res, `Saving to ${def.title}`);
          const created = await res.json();
          snap.set(key, { itemId: created.Id, etag: created['@odata.etag'] || created['odata.etag'] || '*', json });
        } else {
          const res = await this.http.post(this.listApi(def, `/items(${before.itemId})`), JSON.stringify(body), { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': before.etag });
          if (!res.ok) await this.fail(res, `Updating ${def.title}`);
          snap.set(key, { itemId: before.itemId, etag: res.headers.get('ETag') || '*', json });
        }
      }
      for (const [key, s] of [...snap]) {
        if (seen.has(key)) continue;
        const res = await this.http.post(this.listApi(def, `/items(${s.itemId})`), undefined, { 'X-HTTP-Method': 'DELETE', 'IF-MATCH': '*' });
        if (!res.ok && res.status !== 404) await this.fail(res, `Deleting from ${def.title}`);
        snap.delete(key);
      }
    }
  }

  /** Uploads a file to the attachments library and returns a link the hub can store. */
  public async uploadFile(file: File, id: string): Promise<{ name: string; data: string }> {
    // eslint-disable-next-line no-control-regex
    const safe = file.name.replace(/[\\/:*?"<>|#%&{}~'\u0000-\u001f]+/g, '_').replace(/^[.\s]+|[.\s]+$/g, '') || 'file';
    const name = `${new Date().toISOString().slice(0, 10)}_${id.slice(0, 8)}_${safe}`.slice(0, 200);
    const folder = this.listPath(ATTACHMENTS, true);
    const res = await this.http.post(
      `${this.webUrl}/_api/web/GetFolderByServerRelativeUrl(@f)/Files/add(url=@n,overwrite=false)?@f=${odataString(folder)}&@n=${odataString(name)}`,
      await file.arrayBuffer(),
      { 'Content-Type': 'application/octet-stream' }
    );
    if (!res.ok) await this.fail(res, 'Uploading the attachment');
    const saved = await res.json();
    return { name: file.name, data: encodeURI(saved.ServerRelativeUrl) };
  }
}
