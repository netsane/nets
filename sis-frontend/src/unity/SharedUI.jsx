import React, { createContext, useContext, useState, useMemo, useEffect } from 'react';

const UIContext = createContext(null);

export function useUI() {
  const ctx = useContext(UIContext);
  if (!ctx) throw new Error('useUI must be used within UIProvider');
  return ctx;
}

export function UIProvider({ children }) {
  const [backgroundImage, setBackgroundImage] = useState(
    () => (typeof localStorage !== 'undefined' && localStorage.getItem('bgImage')) || null
  );
  const [logoURL, setLogoURL] = useState(
    () => (typeof localStorage !== 'undefined' && localStorage.getItem('logoURL')) || null
  );
  const [primaryColor, setPrimaryColor] = useState(
    () => (typeof localStorage !== 'undefined' && localStorage.getItem('bgColor')) || '#06b6d4'
  );
  const [textColor, setTextColor] = useState(
    () => (typeof localStorage !== 'undefined' && localStorage.getItem('textColor')) || '#ffffff'
  );
  const [bgBlur, setBgBlur] = useState(0);
  const [bgScale, setBgScale] = useState(1);
  const [bgBrightness, setBgBrightness] = useState(100);

  const updateUISettings = (settings = {}) => {
    if (settings.backgroundImage !== undefined) {
      setBackgroundImage(settings.backgroundImage);
      try { localStorage.setItem('bgImage', settings.backgroundImage || ''); } catch (_) {}
    }
    if (settings.logoURL !== undefined) {
      setLogoURL(settings.logoURL);
      try { localStorage.setItem('logoURL', settings.logoURL || ''); } catch (_) {}
    }
    if (settings.primaryColor) {
      setPrimaryColor(settings.primaryColor);
      try { localStorage.setItem('bgColor', settings.primaryColor); } catch (_) {}
    }
    if (settings.textColor) {
      setTextColor(settings.textColor);
      try { localStorage.setItem('textColor', settings.textColor); } catch (_) {}
    }
    if (settings.bgBlur !== undefined) setBgBlur(settings.bgBlur);
    if (settings.bgScale !== undefined) setBgScale(settings.bgScale);
    if (settings.bgBrightness !== undefined) setBgBrightness(settings.bgBrightness);
  };

  const resetToDefault = () => {
    try { localStorage.removeItem('bgImage'); localStorage.removeItem('logoURL'); } catch (_) {}
    setBackgroundImage(null);
    setLogoURL(null);
  };

  /** Standard logo box — use in headers / login */
  const LogoMark = ({ size = 32, className = '' }) => {
    if (!logoURL) {
      return (
        <div
          className={'rounded-lg bg-cyan-600/80 flex items-center justify-center font-black text-white shrink-0 ' + className}
          style={{ width: size, height: size, fontSize: Math.max(10, size * 0.35) }}
        >
          UU
        </div>
      );
    }
    return (
      <img
        src={logoURL}
        alt="Logo"
        className={'object-contain shrink-0 ' + className}
        style={{ width: size, height: size, maxWidth: size, maxHeight: size }}
      />
    );
  };

  const value = useMemo(
    () => ({
      updateUISettings,
      resetToDefault,
      glassOpacityClass: 'bg-gray-900/70 backdrop-blur-md',
      LogoMark,
      uiSettings: {
        backgroundImage,
        logoURL,
        primaryColor,
        textColor,
        bgBlur,
        bgScale,
        bgBrightness,
      },
      rootStyle: backgroundImage
        ? {
            backgroundImage: 'url(' + backgroundImage + ')',
            backgroundSize: 'cover',
            backgroundPosition: 'center',
            position: 'fixed',
            inset: 0,
            zIndex: -1,
            filter: 'brightness(' + bgBrightness + '%) blur(' + bgBlur + 'px)',
            transform: 'scale(' + bgScale + ')',
          }
        : { display: 'none' },
    }),
    [backgroundImage, logoURL, primaryColor, textColor, bgBlur, bgScale, bgBrightness]
  );

  return (
    <UIContext.Provider value={value}>
      {value.rootStyle.display !== 'none' && <div style={value.rootStyle} />}
      <div className="min-h-0 flex-1 flex flex-col" style={{ color: textColor }}>
        {children}
      </div>
    </UIContext.Provider>
  );
}

export default UIProvider;
