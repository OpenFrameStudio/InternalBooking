import { mkdir, readFile, rename, stat, writeFile } from "node:fs/promises";
import path from "node:path";

export function createStorage({ dataDir, dataFiles, storageBackend, supabaseStorage }) {
  async function readJsonFile(filePath, fallback) {
    try {
      const raw = await readFile(filePath, "utf8");
      return JSON.parse(raw);
    } catch (error) {
      if (error.code === "ENOENT") {
        return fallback;
      }
      throw error;
    }
  }

  async function writeJsonFile(filePath, value) {
    await mkdir(path.dirname(filePath), { recursive: true });
    const tempFile = `${filePath}.${process.pid}.${Date.now()}.tmp`;
    await writeFile(tempFile, `${JSON.stringify(value, null, 2)}\n`, "utf8");
    await rename(tempFile, filePath);
  }

  function supabaseTableUrl(query = "") {
    const tableName = encodeURIComponent(supabaseStorage.table);
    return `${supabaseStorage.url}/rest/v1/${tableName}${query}`;
  }

  function supabaseHeaders(extra = {}) {
    return {
      apikey: supabaseStorage.key,
      Authorization: `Bearer ${supabaseStorage.key}`,
      "Content-Type": "application/json",
      ...extra
    };
  }

  async function parseSupabaseError(response) {
    const data = await response.json().catch(() => ({}));
    return data.message || data.details || `Supabase storage request failed with ${response.status}.`;
  }

  function supabaseKey(dataFile) {
    return dataFile.supabaseKey || path.basename(dataFile.file, ".json");
  }

  async function readSupabaseJson(key, fallback) {
    const response = await fetch(supabaseTableUrl(`?namespace=eq.${encodeURIComponent(key)}&select=payload&limit=1`), {
      headers: supabaseHeaders()
    });

    if (!response.ok) {
      throw new Error(await parseSupabaseError(response));
    }

    const rows = await response.json();
    return rows[0]?.payload ?? fallback;
  }

  async function writeSupabaseJson(key, value) {
    const response = await fetch(supabaseTableUrl("?on_conflict=namespace"), {
      method: "POST",
      headers: supabaseHeaders({ Prefer: "resolution=merge-duplicates,return=minimal" }),
      body: JSON.stringify([{
        namespace: key,
        payload: value,
        updated_at: new Date().toISOString()
      }])
    });

    if (!response.ok) {
      throw new Error(await parseSupabaseError(response));
    }
  }

  async function seedSupabaseDataFile(dataFile) {
    const key = supabaseKey(dataFile);
    const existing = await readSupabaseJson(key, undefined);
    if (existing !== undefined) {
      return;
    }

    const seed = await readJsonFile(dataFile.seedFile, dataFile.fallback);
    await writeSupabaseJson(key, seed);
  }

  async function fileExists(filePath) {
    try {
      await stat(filePath);
      return true;
    } catch (error) {
      if (error.code === "ENOENT") {
        return false;
      }
      throw error;
    }
  }

  async function seedDataFile(targetFile, seedFile, fallback) {
    if (await fileExists(targetFile)) {
      return;
    }

    const seed = await readJsonFile(seedFile, fallback);
    await writeJsonFile(targetFile, seed);
  }

  async function readStoredJson(dataFile) {
    if (storageBackend === "supabase") {
      return readSupabaseJson(supabaseKey(dataFile), dataFile.fallback);
    }

    return readJsonFile(dataFile.file, dataFile.fallback);
  }

  async function writeStoredJson(dataFile, value) {
    if (storageBackend === "supabase") {
      await writeSupabaseJson(supabaseKey(dataFile), value);
      return;
    }

    await writeJsonFile(dataFile.file, value);
  }

  async function seedStoredDataFile(dataFile) {
    if (storageBackend === "supabase") {
      await seedSupabaseDataFile(dataFile);
      return;
    }

    await seedDataFile(dataFile.file, dataFile.seedFile, dataFile.fallback);
  }

  async function prepareDataStorage() {
    if (storageBackend !== "supabase") {
      await mkdir(dataDir, { recursive: true });
    }

    await Promise.all(Object.values(dataFiles).map(seedStoredDataFile));
  }

  return {
    prepareDataStorage,
    readStoredJson,
    writeStoredJson
  };
}
