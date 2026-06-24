
import { initializeApp } from 'firebase/app';
import { getAuth, GoogleAuthProvider, signInWithPopup, signInAnonymously } from 'firebase/auth';
import { 
  initializeFirestore, 
  persistentLocalCache, 
  persistentMultipleTabManager, 
  doc, 
  getDocFromServer 
} from 'firebase/firestore';
import firebaseConfig from './firebase-applet-config.json';

// Initialize Firebase
const app = initializeApp(firebaseConfig);

// Initialize Firestore with persistent offline local cache and multi-tab sync
export const db = initializeFirestore(app, {
  localCache: persistentLocalCache({
    tabManager: persistentMultipleTabManager()
  })
}, (firebaseConfig as any).firestoreDatabaseId);

export const auth = getAuth(app);
export const googleProvider = new GoogleAuthProvider();
googleProvider.addScope('https://www.googleapis.com/auth/spreadsheets');

// Cache the access token in memory.
let cachedAccessToken: string | null = null;

export const getCachedAccessToken = () => cachedAccessToken;

export const setCachedAccessToken = (token: string | null) => {
  cachedAccessToken = token;
};

// Test connection
async function testConnection() {
  try {
    await getDocFromServer(doc(db, 'test', 'connection'));
  } catch (error) {
    if (error instanceof Error && error.message.includes('the client is offline')) {
      console.error("Please check your Firebase configuration. The client is offline.");
    }
  }
}
testConnection();

export const loginWithGoogle = async () => {
  try {
    const result = await signInWithPopup(auth, googleProvider);
    const credential = GoogleAuthProvider.credentialFromResult(result);
    if (credential?.accessToken) {
      setCachedAccessToken(credential.accessToken);
    }
    return { user: result.user, error: null };
  } catch (error: any) {
    console.error("Google Login Error:", error);
    return { user: null, error: error.code || error.message };
  }
};

export const loginAnonymously = async (retries = 3, delayMs = 1500) => {
  let lastError: any = null;
  for (let attempt = 1; attempt <= retries; attempt++) {
    try {
      const result = await signInAnonymously(auth);
      return { user: result.user, error: null, isConfigError: false };
    } catch (error: any) {
      lastError = error;
      const isNetworkError = error?.code === 'auth/network-request-failed' || 
                             error?.message?.includes('network-request-failed') ||
                             error?.message?.includes('network connection') ||
                             error?.message?.includes('failed to fetch');
      
      if (isNetworkError && attempt < retries) {
        console.warn(`Anonymous login failed due to network error. Retrying attempt ${attempt}/${retries} in ${delayMs}ms...`, error);
        await new Promise((resolve) => setTimeout(resolve, delayMs));
        delayMs *= 2; // exponential backoff
        continue;
      }
      
      if (error?.code === 'auth/admin-restricted-operation') {
        return { user: null, error, isConfigError: true };
      }
      
      break;
    }
  }
  
  console.error("Anonymous Login Error after retries:", lastError);
  return { user: null, error: lastError, isConfigError: false };
};
