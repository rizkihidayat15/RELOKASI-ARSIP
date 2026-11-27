import React from 'react';
import { AlertTriangle, RefreshCw } from 'lucide-react';

interface ConfirmDialogProps {
  open: boolean;
  title: string;
  description: string;
  confirmLabel: string;
  cancelLabel?: string;
  variant?: 'danger' | 'primary';
  busy?: boolean;
  onConfirm: () => void;
  onCancel: () => void;
}

export const ConfirmDialog: React.FC<ConfirmDialogProps> = ({
  open,
  title,
  description,
  confirmLabel,
  cancelLabel = 'Batalkan',
  variant = 'danger',
  busy = false,
  onConfirm,
  onCancel,
}) => {
  if (!open) return null;

  const isPrimary = variant === 'primary';

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/45 p-4 backdrop-blur-[2px]"
      role="presentation"
      onClick={(e) => {
        if (e.target === e.currentTarget && !busy) onCancel();
      }}
    >
      <div
        role="dialog"
        aria-modal="true"
        aria-labelledby="confirm-dialog-title"
        className="w-full max-w-md rounded-xl border border-slate-200 bg-white p-6 shadow-2xl"
      >
        <div className="flex items-start gap-3">
          <div
            className={`flex h-9 w-9 shrink-0 items-center justify-center rounded-lg ${
              isPrimary
                ? 'bg-emerald-50 text-emerald-700'
                : 'bg-red-50 text-red-600'
            }`}
          >
            {isPrimary ? (
              <RefreshCw className="h-5 w-5" />
            ) : (
              <AlertTriangle className="h-5 w-5" />
            )}
          </div>
          <div>
            <h2
              id="confirm-dialog-title"
              className="text-base font-bold text-slate-900"
            >
              {title}
            </h2>
            <p className="mt-1.5 text-sm leading-relaxed text-slate-600">
              {description}
            </p>
          </div>
        </div>

        <div className="mt-6 flex items-center justify-end gap-3">
          <button
            type="button"
            disabled={busy}
            onClick={onCancel}
            className="rounded-lg border border-slate-300 bg-white px-4 py-2 text-sm font-medium text-slate-700 transition-colors hover:bg-slate-50 disabled:opacity-50 whitespace-nowrap"
          >
            {cancelLabel}
          </button>
          <button
            type="button"
            disabled={busy}
            onClick={onConfirm}
            data-testid="confirm-dialog-submit"
            className={`rounded-lg px-4 py-2 text-sm font-semibold text-white transition-colors disabled:opacity-50 whitespace-nowrap ${
              isPrimary
                ? 'bg-emerald-700 hover:bg-emerald-800'
                : 'bg-red-600 hover:bg-red-700'
            }`}
          >
            {busy ? 'Memproses...' : confirmLabel}
          </button>
        </div>
      </div>
    </div>
  );
};
