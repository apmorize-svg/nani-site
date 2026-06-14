"use client";

import { useState, useCallback, useMemo } from "react";
import * as XLSX from "xlsx";

type Row = Record<string, string>;
type SkuMap = Record<string, string>;

function parseExcel(file: File): Promise<Row[]> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const wb = XLSX.read(data, { type: "array", cellDates: true });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json<Row>(ws, {
          defval: "",
          raw: false,
          dateNF: "dd/mm/yyyy hh:mm:ss",
        });
        resolve(rows);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

function parseSkuMap(file: File): Promise<SkuMap> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const wb = XLSX.read(data, { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json<Row>(ws, { defval: "" });
        const map: SkuMap = {};
        rows.forEach((row) => {
          const vals = Object.values(row);
          if (vals.length >= 2) {
            map[String(vals[0]).trim()] = String(vals[1]).trim();
          }
        });
        resolve(map);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

function extractDate(val: string): string {
  if (!val) return "";
  const match = val.match(/(\d{2})\/(\d{2})\/(\d{4})/);
  if (match) return `${match[3]}-${match[2]}-${match[1]}`;
  return val.slice(0, 10);
}

const DATE_COL = "Data pedido de aviso";
const SKU_COL = "SKU";

export default function Home() {
  const [rows, setRows] = useState<Row[]>([]);
  const [columns, setColumns] = useState<string[]>([]);
  const [skuMap, setSkuMap] = useState<SkuMap>({});
  const [dateFrom, setDateFrom] = useState("");
  const [dateTo, setDateTo] = useState("");
  const [copied, setCopied] = useState(false);
  const [loadingMain, setLoadingMain] = useState(false);
  const [loadingSku, setLoadingSku] = useState(false);

  const handleMainFile = useCallback(async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setLoadingMain(true);
    try {
      const data = await parseExcel(file);
      setRows(data);
      if (data.length > 0) setColumns(Object.keys(data[0]));
    } finally {
      setLoadingMain(false);
    }
  }, []);

  const handleSkuFile = useCallback(async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setLoadingSku(true);
    try {
      const map = await parseSkuMap(file);
      setSkuMap(map);
    } finally {
      setLoadingSku(false);
    }
  }, []);

  const filtered = useMemo(() => {
    return rows.filter((row) => {
      const isoDate = extractDate(row[DATE_COL] || "");
      if (dateFrom && isoDate < dateFrom) return false;
      if (dateTo && isoDate > dateTo) return false;
      return true;
    });
  }, [rows, dateFrom, dateTo]);

  const displayCols = useMemo(() => {
    if (columns.length === 0) return [];
    const skuIdx = columns.indexOf(SKU_COL);
    if (skuIdx === -1) return [...columns, "Nome do Produto"];
    return [
      ...columns.slice(0, skuIdx + 1),
      "Nome do Produto",
      ...columns.slice(skuIdx + 1),
    ];
  }, [columns]);

  const handleCopy = useCallback(async () => {
    const header = displayCols.join("\t");
    const body = filtered
      .map((row) =>
        displayCols
          .map((col) => {
            if (col === "Nome do Produto") return skuMap[String(row[SKU_COL]).trim()] || "";
            return row[col] ?? "";
          })
          .join("\t")
      )
      .join("\n");
    await navigator.clipboard.writeText(header + "\n" + body);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  }, [filtered, displayCols, skuMap]);

  return (
    <main className="min-h-screen bg-gray-50 p-6">
      <div className="max-w-7xl mx-auto space-y-6">
        <h1 className="text-2xl font-bold text-gray-800">Importador de Planilha</h1>

        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
          <div className="bg-white rounded-xl border border-gray-200 p-5 space-y-2">
            <p className="font-semibold text-gray-700">Planilha de pedidos</p>
            <p className="text-sm text-gray-500">Colunas: Email, Telefone, SKU, Data pedido de aviso</p>
            <input
              type="file"
              accept=".xlsx,.xls,.csv"
              onChange={handleMainFile}
              className="block w-full text-sm text-gray-600 file:mr-3 file:py-2 file:px-4 file:rounded-lg file:border-0 file:bg-blue-50 file:text-blue-700 file:font-medium hover:file:bg-blue-100 cursor-pointer"
            />
            {loadingMain && <p className="text-sm text-blue-500">Carregando...</p>}
            {rows.length > 0 && !loadingMain && (
              <p className="text-sm text-blue-600">{rows.length} linhas carregadas</p>
            )}
          </div>

          <div className="bg-white rounded-xl border border-gray-200 p-5 space-y-2">
            <p className="font-semibold text-gray-700">Mapeamento SKU → Nome do produto</p>
            <p className="text-sm text-gray-500">Planilha com duas colunas: SKU | Nome (opcional)</p>
            <input
              type="file"
              accept=".xlsx,.xls,.csv"
              onChange={handleSkuFile}
              className="block w-full text-sm text-gray-600 file:mr-3 file:py-2 file:px-4 file:rounded-lg file:border-0 file:bg-green-50 file:text-green-700 file:font-medium hover:file:bg-green-100 cursor-pointer"
            />
            {loadingSku && <p className="text-sm text-green-500">Carregando...</p>}
            {Object.keys(skuMap).length > 0 && (
              <p className="text-sm text-green-600">{Object.keys(skuMap).length} SKUs mapeados</p>
            )}
          </div>
        </div>

        {rows.length > 0 && (
          <div className="bg-white rounded-xl border border-gray-200 p-5 flex flex-wrap items-end gap-4">
            <div className="space-y-1">
              <label className="text-sm font-medium text-gray-700">Data de</label>
              <input
                type="date"
                value={dateFrom}
                onChange={(e) => setDateFrom(e.target.value)}
                className="block border border-gray-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-400"
              />
            </div>
            <div className="space-y-1">
              <label className="text-sm font-medium text-gray-700">Data até</label>
              <input
                type="date"
                value={dateTo}
                onChange={(e) => setDateTo(e.target.value)}
                className="block border border-gray-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-400"
              />
            </div>
            <button
              onClick={() => { setDateFrom(""); setDateTo(""); }}
              className="px-4 py-2 text-sm rounded-lg border border-gray-300 text-gray-600 hover:bg-gray-100"
            >
              Limpar filtro
            </button>

            <div className="ml-auto flex items-center gap-3">
              <span className="text-sm text-gray-500">{filtered.length} linha{filtered.length !== 1 ? "s" : ""}</span>
              <button
                onClick={handleCopy}
                className="px-5 py-2 rounded-lg bg-blue-600 text-white text-sm font-medium hover:bg-blue-700 active:scale-95 transition-all"
              >
                {copied ? "✓ Copiado!" : "Copiar tabela"}
              </button>
            </div>
          </div>
        )}

        {rows.length > 0 && (
          <div className="bg-white rounded-xl border border-gray-200 overflow-auto">
            <table className="min-w-full text-sm">
              <thead>
                <tr className="bg-gray-100 border-b border-gray-200">
                  {displayCols.map((col) => (
                    <th
                      key={col}
                      className={`px-4 py-3 text-left font-semibold whitespace-nowrap ${
                        col === "Nome do Produto" ? "text-green-700 bg-green-50" : "text-gray-700"
                      }`}
                    >
                      {col}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {filtered.map((row, i) => (
                  <tr key={i} className="border-b border-gray-100 hover:bg-gray-50">
                    {displayCols.map((col) => {
                      const val =
                        col === "Nome do Produto"
                          ? skuMap[String(row[SKU_COL]).trim()] || ""
                          : (row[col] ?? "");
                      return (
                        <td
                          key={col}
                          className={`px-4 py-2 whitespace-nowrap ${
                            col === "Nome do Produto" ? "text-green-700 font-medium" : "text-gray-700"
                          }`}
                        >
                          {val}
                        </td>
                      );
                    })}
                  </tr>
                ))}
              </tbody>
            </table>
            {filtered.length === 0 && (
              <p className="text-center py-8 text-gray-400">Nenhum resultado para o filtro selecionado.</p>
            )}
          </div>
        )}

        {rows.length === 0 && (
          <div className="text-center py-20 text-gray-400">
            <p className="text-lg">Importe uma planilha para começar</p>
          </div>
        )}
      </div>
    </main>
  );
}
