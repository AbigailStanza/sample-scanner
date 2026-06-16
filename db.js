/* db.js — IndexedDB 래퍼
 * 데이터셋(원본 파일 바이트 + 파싱된 행 + 체크/비고 상태)을 저장/복구한다.
 * localStorage(5MB) 한계를 피하고 다중 데이터셋·원본 바이트를 보존하기 위해 IndexedDB 사용.
 *
 * 데이터셋 레코드 형태:
 *   {
 *     id: string,            // dateLabel 기반 키 (예: "2026-06-12") — 같은 날짜 재업로드 시 덮어씀
 *     dateLabel: string,
 *     manager: string,
 *     fileName: string,
 *     sheetName: string,
 *     headerRow: number,     // 원본 시트 0-based 헤더 행
 *     colMap: object,        // {id, vendor, type, name, judge, ...} -> 0-based 컬럼 인덱스
 *     checkColRef: string,   // '확인' 열 글자 (예: 'H')
 *     remarkColRef: string,  // '비고' 열 글자 (예: 'J')
 *     rows: Array<RowState>, // 파싱 결과 + checked/remark
 *     rawBytes: ArrayBuffer, // 원본 .xlsx (내보내기 시 재사용)
 *     updatedAt: number
 *   }
 */
const DB = (() => {
  const DB_NAME = 'sample-scanner';
  const DB_VERSION = 1;
  const STORE = 'datasets';
  const META = 'meta'; // 활성 데이터셋 id 등 보관

  let _db = null;

  function open() {
    if (_db) return Promise.resolve(_db);
    return new Promise((resolve, reject) => {
      const req = indexedDB.open(DB_NAME, DB_VERSION);
      req.onupgradeneeded = () => {
        const db = req.result;
        if (!db.objectStoreNames.contains(STORE)) {
          db.createObjectStore(STORE, { keyPath: 'id' });
        }
        if (!db.objectStoreNames.contains(META)) {
          db.createObjectStore(META, { keyPath: 'key' });
        }
      };
      req.onsuccess = () => { _db = req.result; resolve(_db); };
      req.onerror = () => reject(req.error);
    });
  }

  function tx(storeNames, mode) {
    return open().then((db) => db.transaction(storeNames, mode));
  }

  function reqPromise(request) {
    return new Promise((resolve, reject) => {
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
  }

  return {
    /** 데이터셋 저장(또는 덮어쓰기) */
    async saveDataset(ds) {
      ds.updatedAt = Date.now();
      const t = await tx([STORE], 'readwrite');
      await reqPromise(t.objectStore(STORE).put(ds));
      return ds.id;
    },

    /** 단일 행 상태만 갱신(autosave 디바운스용) */
    async updateRowState(id, rowIndexInArray, patch) {
      const t = await tx([STORE], 'readwrite');
      const store = t.objectStore(STORE);
      const ds = await reqPromise(store.get(id));
      if (!ds) return;
      Object.assign(ds.rows[rowIndexInArray], patch);
      ds.updatedAt = Date.now();
      await reqPromise(store.put(ds));
    },

    /** 데이터셋 전체 행 상태 일괄 저장 */
    async saveRows(id, rows) {
      const t = await tx([STORE], 'readwrite');
      const store = t.objectStore(STORE);
      const ds = await reqPromise(store.get(id));
      if (!ds) return;
      ds.rows = rows;
      ds.updatedAt = Date.now();
      await reqPromise(store.put(ds));
    },

    async getDataset(id) {
      const t = await tx([STORE], 'readonly');
      return reqPromise(t.objectStore(STORE).get(id));
    },

    async listDatasets() {
      const t = await tx([STORE], 'readonly');
      const all = await reqPromise(t.objectStore(STORE).getAll());
      // 날짜 내림차순(최신 먼저)
      return all.sort((a, b) => (b.dateLabel || '').localeCompare(a.dateLabel || ''));
    },

    async deleteDataset(id) {
      const t = await tx([STORE], 'readwrite');
      await reqPromise(t.objectStore(STORE).delete(id));
    },

    async setActiveId(id) {
      const t = await tx([META], 'readwrite');
      await reqPromise(t.objectStore(META).put({ key: 'activeId', value: id }));
    },

    async getActiveId() {
      const t = await tx([META], 'readonly');
      const r = await reqPromise(t.objectStore(META).get('activeId'));
      return r ? r.value : null;
    },
  };
})();
