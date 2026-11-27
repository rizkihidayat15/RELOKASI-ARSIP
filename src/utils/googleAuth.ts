import { getApps, initializeApp } from 'firebase/app';
import {
  getAuth,
  GoogleAuthProvider,
  onAuthStateChanged,
  signInWithPopup,
  type User,
} from 'firebase/auth';
import firebaseConfig from '../../firebase-applet-config.json';

export const WORKSPACE_SCOPES = [
  'https://www.googleapis.com/auth/spreadsheets',
  'https://www.googleapis.com/auth/drive.metadata.readonly',
] as const;

const app = getApps().length > 0 ? getApps()[0] : initializeApp(firebaseConfig);
const auth = getAuth(app);

const provider = new GoogleAuthProvider();
for (const scope of WORKSPACE_SCOPES) {
  provider.addScope(scope);
}
provider.setCustomParameters({ prompt: 'select_account' });

let isSigningIn = false;
// Keep OAuth access token strictly in memory — never persist to localStorage/sessionStorage.
let cachedAccessToken: string | null = null;

export const initGoogleAuth = (
  onAuthSuccess?: (user: User, token: string) => void,
  onAuthRequired?: (user: User | null) => void
) => {
  return onAuthStateChanged(auth, (user: User | null) => {
    if (user) {
      if (cachedAccessToken) {
        onAuthSuccess?.(user, cachedAccessToken);
      } else if (!isSigningIn) {
        cachedAccessToken = null;
        onAuthRequired?.(user);
      }
    } else {
      cachedAccessToken = null;
      onAuthRequired?.(null);
    }
  });
};

export const signInWithGoogleWorkspace = async (): Promise<{
  user: User;
  accessToken: string;
}> => {
  try {
    isSigningIn = true;
    const result = await signInWithPopup(auth, provider);
    const credential = GoogleAuthProvider.credentialFromResult(result);
    if (!credential?.accessToken) {
      throw new Error(
        'Gagal mendapatkan akses Google Sheets & Drive. Silakan coba masuk kembali.'
      );
    }
    cachedAccessToken = credential.accessToken;
    return {
      user: result.user,
      accessToken: cachedAccessToken,
    };
  } finally {
    isSigningIn = false;
  }
};

export const getWorkspaceAccessToken = (): string | null => cachedAccessToken;

export const clearWorkspaceAccessToken = (): void => {
  cachedAccessToken = null;
};

export const logoutGoogleWorkspace = async (): Promise<void> => {
  cachedAccessToken = null;
  await auth.signOut();
};
