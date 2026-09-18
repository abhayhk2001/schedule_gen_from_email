const OVERRIDE_KEY = "addCalEvent.themeOverride";

function safeStorageGet(key) {
  try {
    return window.localStorage?.getItem(key) ?? null;
  } catch {
    return null;
  }
}

function safeStorageSet(key, value) {
  try {
    if (value === null) window.localStorage.removeItem(key);
    else window.localStorage.setItem(key, value);
  } catch {}
}

function hexToRgb(hex) {
  if (typeof hex !== "string") return { r: 0, g: 0, b: 0 };
  const trimmed = hex.trim().replace(/^#/, "");
  const value =
    trimmed.length === 3
      ? trimmed
          .split("")
          .map((c) => c + c)
          .join("")
      : trimmed;
  if (!/^[0-9a-fA-F]{6}$/.test(value)) return { r: 0, g: 0, b: 0 };
  const intVal = parseInt(value, 16);
  return {
    r: (intVal >> 16) & 0xff,
    g: (intVal >> 8) & 0xff,
    b: intVal & 0xff,
  };
}

function relativeLuminance({ r, g, b }) {
  const channel = (c) => {
    const s = c / 255;
    return s <= 0.03928 ? s / 12.92 : Math.pow((s + 0.055) / 1.055, 2.4);
  };
  return 0.2126 * channel(r) + 0.7152 * channel(g) + 0.0722 * channel(b);
}

function isDarkBackground(hex) {
  return relativeLuminance(hexToRgb(hex)) < 0.5;
}

function setToken(name, value) {
  document.documentElement.style.setProperty(name, value);
}

function clearToken(name) {
  document.documentElement.style.removeProperty(name);
}

function applyTokens(palette) {
  const root = document.documentElement;
  setToken("--bg", palette.bg);
  setToken("--fg", palette.fg);
  setToken("--surface", palette.surface);
  setToken("--surface-fg", palette.surfaceFg);
  setToken("--border", palette.border);
  setToken("--accent", palette.accent);

  const accentRgb = hexToRgb(palette.accent);
  const hover = `rgb(${Math.max(
    0,
    Math.round(accentRgb.r * 0.85),
  )}, ${Math.max(0, Math.round(accentRgb.g * 0.85))}, ${Math.max(
    0,
    Math.round(accentRgb.b * 0.85),
  )})`;
  setToken("--accent-hover", hover);
  setToken(
    "--accent-on",
    relativeLuminance(accentRgb) < 0.5 ? "#ffffff" : "#000000",
  );
  setToken(
    "--accent-soft",
    `rgba(${accentRgb.r}, ${accentRgb.g}, ${accentRgb.b}, 0.18)`,
  );

  const dark = isDarkBackground(palette.bg);
  root.dataset.theme = dark ? "dark" : "light";
  root.style.colorScheme = dark ? "dark" : "light";
}

function clearRuntimeTokens() {
  for (const name of [
    "--bg",
    "--fg",
    "--surface",
    "--surface-fg",
    "--border",
    "--accent",
    "--accent-hover",
    "--accent-on",
    "--accent-soft",
  ]) {
    clearToken(name);
  }
}

function applyOfficeTheme() {
  const theme = Office?.context?.officeTheme;
  if (!theme) return false;
  applyTokens({
    bg: theme.bodyBackgroundColor,
    fg: theme.bodyForegroundColor,
    surface: theme.controlBackgroundColor,
    surfaceFg: theme.controlForegroundColor,
    border: theme.controlBorderColor,
    accent: theme.accent1,
  });
  return true;
}

function applyStaticMode(mode) {
  if (mode !== "light" && mode !== "dark") return;
  clearRuntimeTokens();
  document.documentElement.dataset.staticTheme = mode;
  document.documentElement.style.colorScheme = mode;
}

function readOverride() {
  const value = safeStorageGet(OVERRIDE_KEY);
  return value === "light" || value === "dark" ? value : null;
}

export function effectiveMode() {
  const override = readOverride();
  if (override) return override;
  const theme = Office?.context?.officeTheme;
  if (theme) return isDarkBackground(theme.bodyBackgroundColor) ? "dark" : "light";
  return "dark";
}

export function setThemeOverride(mode) {
  if (mode !== "light" && mode !== "dark" && mode !== null) return;
  if (mode === null) {
    safeStorageSet(OVERRIDE_KEY, null);
  } else {
    safeStorageSet(OVERRIDE_KEY, mode);
  }
  if (mode === "light" || mode === "dark") {
    applyStaticMode(mode);
  } else if (Office?.context?.officeTheme) {
    applyOfficeTheme();
  } else {
    applyStaticMode("dark");
  }
}

function subscribeToOfficeThemeChanges() {
  if (Office?.context?.addHandlerAsync) {
    try {
      Office.context.addHandlerAsync(
        Office.EventType.ThemeChanged,
        () => {
          if (!readOverride() && Office?.context?.officeTheme) {
            applyOfficeTheme();
          }
        },
      );
    } catch {}
  }
}

function subscribeToCrossFrameSync() {
  window.addEventListener("storage", (event) => {
    if (event.key !== OVERRIDE_KEY) return;
    const next = event.newValue === "light" || event.newValue === "dark"
      ? event.newValue
      : null;
    if (next) {
      applyStaticMode(next);
    } else if (Office?.context?.officeTheme) {
      applyOfficeTheme();
    }
  });
}

export function initTheme() {
  const override = readOverride();
  if (override) {
    applyStaticMode(override);
    return;
  }
  if (!applyOfficeTheme()) {
    setTimeout(() => {
      if (!readOverride() && !applyOfficeTheme()) {
        applyStaticMode("dark");
      }
    }, 50);
  }
  subscribeToOfficeThemeChanges();
  subscribeToCrossFrameSync();
}
