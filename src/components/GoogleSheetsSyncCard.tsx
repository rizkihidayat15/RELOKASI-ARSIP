import React, { useEffect, useState } from 'react';
import {
  AlertCircle,
  AlertTriangle,
  CheckCircle2,
  ExternalLink,
  FolderOpen,
  Link2,
  LogOut,
  Plus,
  RefreshCw,
  Unlink,
} from 'lucide-react';
import type {
  ConnectedGoogleSheet,
  DriveSpreadsheetItem,
} from '../utils/googleSheetsSync';
import { GoogleSignInButton } from './GoogleSignInButton';

interface GoogleSheetsSyncCardProps {
  isAuthenticated: boolean;
  userEmail: string | null;
  isSigningIn: boolean;
  connectedSheet: ConnectedGoogleSheet | null;
  driveFiles: DriveSpreadsheetItem[];
  isLoadingDriveFiles: boolean;
  isConnectingSheet: boolean;
  isSyncing: boolean;
  isRefreshingSheet: boolean;
  syncError: string | null;
  savedRecordCount: number;
  pendingSyncCount: number;
  pendingNewCount: number;
  pendingUpdateCount: number;
  onSignIn: () => void;
  onSignOut: () => void;
  onConnectByUrl: (urlOrId: string) => Promise<boolean>;
  onLoadDriveFiles: () => void;
  onCreateNewSheet: () => void;
  onSelectSheetTab: (sheetName: string) => void;
  onDisconnectSheet: () => void;
  onRefreshFromSheet: () => void;
  onRequestSyncToGoogleSheet: () => void;
}

export const GoogleSheetsSyncCard: React.FC<GoogleSheetsSyncCardProps> = ({
  isAuthenticated,
  userEmail,
  isSigningIn,
  connectedSheet,
  driveFiles,
  isLoadingDriveFiles,
  isConnectingSheet,
  isSyncing,
  isRefreshingSheet,
  syncError,
  savedRecordCount,
  pendingSyncCount,
  pendingNewCount,
  pendingUpdateCount,
  onSignIn,
  onSignOut,
  onConnectByUrl,
  onLoadDriveFiles,
  onCreateNewSheet,
  onSelectSheetTab,
  onDisconnectSheet,
  onRefreshFromSheet,
  onRequestSyncToGoogleSheet,
}) => {
  const [urlInput, setUrlInput] = useState(
    connectedSheet?.spreadsheetUrl ?? ''
  );
  const [localError, setLocalError] = useState<string | null>(null);
  const [showDriveList, setShowDriveList] = useState(false);

  useEffect(() => {
    setUrlInput(connectedSheet?.spreadsheetUrl ?? '');
  }, [connectedSheet?.spreadsheetUrl]);

  const handleLinkSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setLocalError(null);
    const trimmed = urlInput.trim();
    if (!trimmed) {
      setLocalError('Tempel link Google Spreadsheet atau masukkan ID file.');
      return;
    }
    const ok = await onConnectByUrl(trimmed);
    if (!ok) {
      setLocalError(
        'Link tidak valid. Gunakan link Google Sheets (https://docs.google.com/spreadsheets/d/...).'
      );
    }
  };

  const handleToggleDrivePicker = () => {
    const next = !showDriveList;
    setShowDriveList(next);
    if (next && driveFiles.length === 0) {
      onLoadDriveFiles();
    }
  };

  const displayError = localError || syncError;
  const needsSessionReauth = Boolean(connectedSheet && !isAuthenticated);
  const hasPendingChanges = Boolean(
    connectedSheet && isAuthenticated && pendingSyncCount > 0
  );

  const summaryCardClasses = needsSessionReauth
    ? 'border-amber-200 bg-amber-50/80'
    : hasPendingChanges
      ? 'border-amber-200 bg-amber-50/50'
      : 'border-emerald-200 bg-emerald-50/70';

  return (
    <div
      data-testid="google-sheets-sync-card"
      className="mt-3.5 rounded-lg border border-slate-200 bg-slate-50/80 p-3.5"
    >
      <div className="flex flex-wrap items-center justify-between gap-2">
        <div className="min-w-0">
          <p className="text-xs font-bold text-slate-900">
            Sinkronisasi Google Sheets (Google Drive)
          </p>
          <p className="mt-0.5 text-[11px] text-slate-500">
            Update atau tambah baris langsung di link spreadsheet yang sama
            tanpa membuat file baru.
          </p>
        </div>
        {isAuthenticated && userEmail && (
          <div className="flex items-center gap-2 text-[11px] text-slate-600">
            <span className="max-w-[160px] truncate font-medium">
              {userEmail}
            </span>
            <button
              type="button"
              onClick={onSignOut}
              title="Keluar dari akun Google"
              className="inline-flex items-center gap-1 rounded border border-slate-200 bg-white px-2 py-1 text-[11px] font-medium text-slate-600 transition-colors hover:bg-slate-100 hover:text-slate-900 whitespace-nowrap shrink-0"
            >
              <LogOut className="h-3 w-3" />
              Keluar
            </button>
          </div>
        )}
      </div>

      {/* Connected Target Summary */}
      {connectedSheet && (
        <div
          data-testid="connected-sheet-summary"
          data-sync-state={
            needsSessionReauth
              ? 'session-inactive'
              : hasPendingChanges
                ? 'pending-push'
                : 'synced'
          }
          className={`mt-3 rounded-lg border p-2.5 text-xs ${summaryCardClasses}`}
        >
          <div className="flex items-start justify-between gap-2">
            <div className="min-w-0 flex-1">
              <div className="flex items-center gap-1.5">
                {needsSessionReauth || hasPendingChanges ? (
                  <AlertTriangle className="h-3.5 w-3.5 shrink-0 text-amber-700" />
                ) : (
                  <CheckCircle2 className="h-3.5 w-3.5 shrink-0 text-emerald-700" />
                )}
                <p
                  className={`truncate font-semibold ${
                    needsSessionReauth || hasPendingChanges
                      ? 'text-amber-950'
                      : 'text-emerald-950'
                  }`}
                >
                  {connectedSheet.title}
                </p>
              </div>
              <a
                href={connectedSheet.spreadsheetUrl}
                target="_blank"
                rel="noopener noreferrer"
                className="mt-1 inline-flex max-w-full items-center gap-1 text-[11px] font-medium text-[#0F4C81] hover:underline"
              >
                <span className="truncate font-mono">
                  docs.google.com/spreadsheets/d/
                  {connectedSheet.spreadsheetId.slice(0, 14)}...
                </span>
                <ExternalLink className="h-3 w-3 shrink-0" />
              </a>
            </div>
            <button
              type="button"
              onClick={onDisconnectSheet}
              aria-label="Putuskan link Google Spreadsheet"
              title="Ganti link spreadsheet"
              className="rounded p-1 text-slate-500 transition-colors hover:bg-white hover:text-red-600 shrink-0"
            >
              <Unlink className="h-3.5 w-3.5" />
            </button>
          </div>

          {/* Explicit Connection / Sync State Message */}
          {needsSessionReauth ? (
            <p
              data-testid="sheets-session-inactive-notice"
              className="mt-2 text-[11px] leading-relaxed font-medium text-amber-900"
            >
              Link tersimpan, tetapi sesi Google belum aktif (perlu login ulang
              setelah halaman dimuat ulang). Klik{' '}
              <strong className="font-bold">Sign in with Google</strong> di
              bawah agar data dapat membaca &amp; mengupdate link ini.
              {savedRecordCount > 0
                ? ` (${savedRecordCount} baris lokal menunggu disinkronkan.)`
                : ''}
            </p>
          ) : hasPendingChanges ? (
            <p
              data-testid="sheets-pending-sync-notice"
              className="mt-2 text-[11px] leading-relaxed font-medium text-amber-900 tabular-nums"
            >
              {pendingSyncCount} perubahan belum dikirim ke Google Sheets (+
              {pendingNewCount} baru, {pendingUpdateCount} diperbarui). Klik
              tombol kirim di bawah untuk memperbarui link spreadsheet.
            </p>
          ) : (
            <p
              data-testid="sheets-synced-notice"
              className="mt-2 text-[11px] font-medium text-emerald-800"
            >
              Tersinkron · Auto-baca dari Google Sheets aktif setiap 30 detik
              &amp; saat kembali ke tab ini.
            </p>
          )}

          {/* Sheet Tab Selector & Refresh */}
          <div className="mt-2 flex flex-wrap items-center justify-between gap-2">
            <div className="flex flex-wrap items-center gap-2">
              <label
                htmlFor="select-google-sheet-tab"
                className={`text-[11px] font-medium ${
                  needsSessionReauth || hasPendingChanges
                    ? 'text-amber-900'
                    : 'text-emerald-900'
                }`}
              >
                Tab Sheet:
              </label>
              <select
                id="select-google-sheet-tab"
                value={connectedSheet.sheetName}
                onChange={(e) => onSelectSheetTab(e.target.value)}
                disabled={!isAuthenticated}
                className="rounded border border-slate-300 bg-white px-2 py-0.5 text-[11px] font-medium text-slate-800 focus:border-[#0F4C81] focus:outline-none disabled:opacity-60"
              >
                {connectedSheet.availableSheets.map((tab) => (
                  <option key={tab} value={tab}>
                    {tab}
                  </option>
                ))}
              </select>
            </div>
            {isAuthenticated && (
              <button
                type="button"
                onClick={onRefreshFromSheet}
                disabled={isRefreshingSheet || isSyncing}
                className="inline-flex items-center gap-1 rounded border border-slate-300 bg-white px-2 py-0.5 text-[11px] font-medium text-slate-700 transition-colors hover:bg-slate-50 disabled:opacity-50 whitespace-nowrap shrink-0"
              >
                <RefreshCw
                  className={`h-3 w-3 ${isRefreshingSheet ? 'animate-spin' : ''}`}
                />
                {isRefreshingSheet ? 'Membaca...' : 'Baca Ulang'}
              </button>
            )}
          </div>

          {connectedSheet.lastSyncSummary && connectedSheet.lastSyncedAt && (
            <p className="mt-2 text-[11px] text-slate-700 tabular-nums">
              Terakhir kirim:{' '}
              {new Date(connectedSheet.lastSyncedAt).toLocaleTimeString(
                'id-ID',
                { hour: '2-digit', minute: '2-digit' }
              )}{' '}
              · +{connectedSheet.lastSyncSummary.appendedCount} baru,{' '}
              {connectedSheet.lastSyncSummary.updatedCount} diperbarui
            </p>
          )}
        </div>
      )}

      {/* Auth Gate OR Link Input */}
      {!isAuthenticated ? (
        <div className="mt-3 space-y-2.5">
          <div className="flex flex-wrap items-center gap-2.5">
            <GoogleSignInButton
              onClick={onSignIn}
              loading={isSigningIn}
              label={
                connectedSheet
                  ? 'Sign in with Google untuk Sinkron'
                  : 'Sign in with Google'
              }
            />
          </div>

          {/* Allow pasting link even before sign-in so target link is ready */}
          {!connectedSheet && (
            <form onSubmit={handleLinkSubmit} className="space-y-1.5 pt-1">
              <label
                htmlFor="input-google-sheet-url"
                className="block text-[11px] font-medium text-slate-600"
              >
                Atau tempel link Google Spreadsheet tujuan lebih dulu:
              </label>
              <div className="flex items-center gap-1.5">
                <input
                  id="input-google-sheet-url"
                  type="text"
                  value={urlInput}
                  onChange={(e) => {
                    setUrlInput(e.target.value);
                    if (localError) setLocalError(null);
                  }}
                  placeholder="https://docs.google.com/spreadsheets/d/..."
                  className="min-w-0 flex-1 rounded-md border border-slate-300 bg-white px-2.5 py-1.5 text-xs text-slate-900 placeholder:text-slate-400 focus:border-[#0F4C81] focus:outline-none"
                />
                <button
                  type="submit"
                  disabled={isConnectingSheet}
                  className="inline-flex items-center gap-1 rounded-md border border-slate-300 bg-white px-2.5 py-1.5 text-xs font-semibold text-slate-700 transition-colors hover:bg-slate-100 disabled:opacity-45 whitespace-nowrap shrink-0"
                >
                  <Link2 className="h-3.5 w-3.5" />
                  Simpan Link
                </button>
              </div>
            </form>
          )}
        </div>
      ) : (
        <div className="mt-3 space-y-2.5">
          {!connectedSheet && (
            <>
              <form onSubmit={handleLinkSubmit} className="space-y-1.5">
                <label
                  htmlFor="input-google-sheet-url"
                  className="block text-[11px] font-semibold text-slate-700"
                >
                  Link Google Spreadsheet Tujuan
                </label>
                <div className="flex items-center gap-1.5">
                  <input
                    id="input-google-sheet-url"
                    type="text"
                    value={urlInput}
                    onChange={(e) => {
                      setUrlInput(e.target.value);
                      if (localError) setLocalError(null);
                    }}
                    placeholder="https://docs.google.com/spreadsheets/d/..."
                    className="min-w-0 flex-1 rounded-md border border-slate-300 bg-white px-2.5 py-1.5 text-xs text-slate-900 placeholder:text-slate-400 focus:border-[#0F4C81] focus:outline-none"
                  />
                  <button
                    type="submit"
                    disabled={isConnectingSheet}
                    className="inline-flex items-center gap-1 rounded-md bg-[#0F4C81] px-3 py-1.5 text-xs font-semibold text-white transition-colors hover:bg-[#0B3A63] disabled:opacity-45 whitespace-nowrap shrink-0"
                  >
                    <Link2 className="h-3.5 w-3.5" />
                    {isConnectingSheet ? 'Memeriksa...' : 'Hubungkan'}
                  </button>
                </div>
              </form>

              <div className="flex flex-wrap items-center gap-2">
                <button
                  type="button"
                  onClick={handleToggleDrivePicker}
                  disabled={isLoadingDriveFiles}
                  className="inline-flex items-center gap-1.5 rounded-md border border-slate-300 bg-white px-2.5 py-1.5 text-[11px] font-semibold text-slate-700 transition-colors hover:bg-slate-100 whitespace-nowrap shrink-0"
                >
                  <FolderOpen className="h-3.5 w-3.5 text-[#0F4C81]" />
                  {isLoadingDriveFiles
                    ? 'Memuat Drive...'
                    : 'Pilih dari Google Drive'}
                </button>
                <button
                  type="button"
                  onClick={onCreateNewSheet}
                  disabled={isConnectingSheet}
                  className="inline-flex items-center gap-1.5 rounded-md border border-slate-300 bg-white px-2.5 py-1.5 text-[11px] font-semibold text-slate-700 transition-colors hover:bg-slate-100 whitespace-nowrap shrink-0"
                >
                  <Plus className="h-3.5 w-3.5 text-emerald-700" />
                  Buat Spreadsheet Baru
                </button>
              </div>

              {showDriveList && (
                <div className="rounded-md border border-slate-200 bg-white">
                  <div className="flex items-center justify-between border-b border-slate-100 px-3 py-2">
                    <span className="text-[11px] font-semibold text-slate-700">
                      Spreadsheet di Google Drive Anda
                    </span>
                    <button
                      type="button"
                      onClick={onLoadDriveFiles}
                      disabled={isLoadingDriveFiles}
                      className="text-[11px] font-medium text-[#0F4C81] hover:underline"
                    >
                      Muat ulang
                    </button>
                  </div>
                  {isLoadingDriveFiles ? (
                    <p className="px-3 py-4 text-center text-xs text-slate-500">
                      Mengambil daftar spreadsheet...
                    </p>
                  ) : driveFiles.length === 0 ? (
                    <p className="px-3 py-4 text-center text-xs text-slate-500">
                      Tidak ditemukan Google Spreadsheet di Drive Anda.
                    </p>
                  ) : (
                    <ul className="max-h-44 divide-y divide-slate-100 overflow-y-auto">
                      {driveFiles.map((file) => (
                        <li key={file.id}>
                          <button
                            type="button"
                            onClick={async () => {
                              const ok = await onConnectByUrl(file.id);
                              if (ok) setShowDriveList(false);
                            }}
                            className="flex w-full items-center justify-between gap-2 px-3 py-2 text-left text-xs transition-colors hover:bg-sky-50"
                          >
                            <span className="truncate font-medium text-slate-800">
                              {file.name}
                            </span>
                            <span className="text-[10px] text-slate-400 whitespace-nowrap shrink-0">
                              Pilih
                            </span>
                          </button>
                        </li>
                      ))}
                    </ul>
                  )}
                </div>
              )}
            </>
          )}

          {connectedSheet && (
            <button
              type="button"
              disabled={savedRecordCount === 0 || isSyncing}
              onClick={onRequestSyncToGoogleSheet}
              data-testid="btn-sync-google-sheets"
              className={`inline-flex w-full items-center justify-center gap-2 rounded-lg px-3.5 py-2 text-xs font-semibold text-white transition-colors disabled:opacity-45 whitespace-nowrap ${
                hasPendingChanges
                  ? 'bg-[#0F4C81] hover:bg-[#0B3A63]'
                  : 'bg-emerald-700 hover:bg-emerald-800'
              }`}
            >
              <RefreshCw
                className={`h-3.5 w-3.5 ${isSyncing ? 'animate-spin' : ''}`}
              />
              {isSyncing
                ? 'Menyinkronkan ke Google Sheets...'
                : hasPendingChanges
                  ? `Kirim ${pendingSyncCount} Perubahan ke Google Sheets`
                  : `Update / Cek Sinkron Google Sheets (${savedRecordCount})`}
            </button>
          )}
        </div>
      )}

      {displayError && (
        <div
          role="alert"
          className="mt-2.5 flex items-start gap-2 rounded-md border border-red-200 bg-red-50 px-2.5 py-2 text-[11px] text-red-700"
        >
          <AlertCircle className="mt-0.5 h-3.5 w-3.5 shrink-0" />
          <span>{displayError}</span>
        </div>
      )}
    </div>
  );
};
