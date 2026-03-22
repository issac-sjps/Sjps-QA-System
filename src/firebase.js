import { initializeApp } from 'firebase/app'
import { getAuth, GoogleAuthProvider } from 'firebase/auth'
import { getFirestore } from 'firebase/firestore'

const firebaseConfig = {
  apiKey:            "AIzaSyC5QDx296J0rhyUCglMcA8sF6AWOMmUWK4",
  authDomain:        "sjps-qa-system.firebaseapp.com",
  projectId:         "sjps-qa-system",
  storageBucket:     "sjps-qa-system.firebasestorage.app",
  messagingSenderId: "27359156834",
  appId:             "1:27359156834:web:324d8a80785dc2c5da8f3b",
  measurementId:     "G-D5D3F7JY4M"
}

const app = initializeApp(firebaseConfig)
export const auth = getAuth(app)
export const db   = getFirestore(app)
export const googleProvider = new GoogleAuthProvider()
