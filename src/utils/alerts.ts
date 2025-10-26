import Swal, { SweetAlertIcon, SweetAlertOptions } from 'sweetalert2';

// Centralized SweetAlert2 wrapper with Sunbeth theme
const base = Swal.mixin({
  buttonsStyling: false,
  // Map to existing button styles
  customClass: {
    popup: 'swal2-sunbeth',
    title: 'swal2-sunbeth-title',
    htmlContainer: 'swal2-sunbeth-body',
    confirmButton: 'btn',
    cancelButton: 'btn ghost',
    denyButton: 'btn ghost',
    actions: 'swal2-sunbeth-actions'
  },
  showClass: {
    popup: 'swal2-animate-in'
  },
  hideClass: {
    popup: 'swal2-animate-out'
  }
});

// Toast preset
const toast = Swal.mixin({
  toast: true,
  position: 'bottom-end',
  timer: 3600,
  timerProgressBar: true,
  showConfirmButton: false,
  showCloseButton: true,
  buttonsStyling: false,
  customClass: {
    popup: 'swal2-sunbeth swal2-toast-sunbeth',
    title: 'swal2-sunbeth-title',
    htmlContainer: 'swal2-sunbeth-body'
  },
  showClass: { popup: 'swal2-animate-in' },
  hideClass: { popup: 'swal2-animate-out' }
});

export async function alertSuccess(title: string, text?: string, opts?: SweetAlertOptions) {
  return base.fire({ icon: 'success', title, html: text, ...opts });
}
export async function alertError(title: string, text?: string, opts?: SweetAlertOptions) {
  return base.fire({ icon: 'error', title, html: text, ...opts });
}
export async function alertInfo(title: string, text?: string, opts?: SweetAlertOptions) {
  return base.fire({ icon: 'info', title, html: text, ...opts });
}
export async function alertWarning(title: string, text?: string, opts?: SweetAlertOptions) {
  return base.fire({ icon: 'warning', title, html: text, ...opts });
}
export async function confirmDialog(
  title: string,
  text?: string,
  confirmText = 'Confirm',
  cancelText = 'Cancel',
  opts?: SweetAlertOptions
): Promise<boolean> {
  const res = await base.fire({
    icon: 'warning',
    title,
    html: text,
    showCancelButton: true,
    confirmButtonText: confirmText,
    cancelButtonText: cancelText,
    ...opts
  });
  return !!res.isConfirmed;
}

// A three-option dialog (confirm/deny/cancel). Useful when you need an extra action like "Preview".
export async function tripleDialog(
  title: string,
  text: string | undefined,
  confirmText: string,
  cancelText: string,
  denyText?: string,
  opts?: SweetAlertOptions
): Promise<'confirm' | 'deny' | 'cancel'> {
  const res = await base.fire({
    icon: 'info',
    title,
    html: text,
    showCancelButton: true,
    showDenyButton: !!denyText,
    confirmButtonText: confirmText,
    cancelButtonText: cancelText,
    denyButtonText: denyText,
    ...opts
  });
  if (res.isConfirmed) return 'confirm';
  if (res.isDenied) return 'deny';
  return 'cancel';
}
export function showToast(title: string, icon: SweetAlertIcon = 'success') {
  return toast.fire({ icon, title });
}
export function showToastHtml(html: string, icon: SweetAlertIcon = 'success') {
  return toast.fire({ icon, html });
}

export default {
  success: alertSuccess,
  error: alertError,
  info: alertInfo,
  warning: alertWarning,
  confirm: confirmDialog,
  tripleDialog,
  toast: showToast,
  toastHtml: showToastHtml
};
