import { CapacitorConfig } from '@capacitor/cli';

const config: CapacitorConfig = {
  appId: 'com.brightmetalwork.portal',
  appName: 'Bright Metalwork',
  webDir: 'dist',
  server: {
    androidScheme: 'https',
    // Remove the line below to use bundled app (offline capable)
    // url: 'https://your-netlify-url.netlify.app'
    cleartext: false
  },
  android: {
    buildOptions: {
      minSdkVersion: 24,          // Android 7.0+ (covers 99%+ of devices)
      targetSdkVersion: 34,       // Android 14
      compileSdkVersion: 34
    },
    allowMixedContent: false,
    captureInput: true,
    webContentsDebuggingEnabled: false
  },
  plugins: {
    Geolocation: {
      // GPS used for site sign-in
    },
    PushNotifications: {
      presentationOptions: ['badge', 'sound', 'alert']
    },
    SplashScreen: {
      launchShowDuration: 2000,
      backgroundColor: '#0a0e1a',
      showSpinner: false
    }
  }
};

export default config;
