(() => {
  const IMAGE_EXTENSIONS = new Set([
    ".jpg",
    ".jpeg",
    ".png",
    ".gif",
    ".bmp",
    ".webp",
    ".tif",
    ".tiff",
    ".jfif",
    ".avif",
    ".heic",
    ".heif",
  ]);
  const RESIZE_MAX = 1920;
  const RESIZE_QUALITY = 0.85;
  const DEFAULT_TAGS = {
    report: ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10"],
    observations: [
      "Absturzsicherung",
      "Arbeiten in der Höhe",
      "Arbeitsmittelprüfung",
      "Beleuchtung",
      "Brandschutz",
      "Chemikalien",
      "Druckbehälter",
      "Elektrik",
      "Ergonomie",
      "Erste Hilfe",
      "Fluchtwege",
      "Gefährdungsbeurteilung",
      "Handwerkzeuge",
      "Kennzeichnung",
      "Lagerung",
      "Lärm",
      "Leitern",
      "Maschinenwartung",
      "Maschinensicherung",
      "Notfallorganisation",
      "PSA",
      "Rutschgefahr",
      "Sauberkeit",
      "Schulung / Unterweisung",
      "Schweißarbeiten",
      "Sicherheitsdatenblätter",
      "Staplerverkehr",
      "Verkehrswege",
      "Werkstatt / Ordnung",
    ],
    training: [
      "Unterlassen",
      "Dulden",
      "Handeln",
      "Vorbild",
      "Iceberg",
      "Pyramide",
      "STOP",
      "SOS",
      "Verhindern",
      "Audit",
      "Risikobeurteilung",
      "AVIVA",
      "StGB Art. 230",
    ],
  };

  const DEMO_PHOTO_URLS = [
    "./demo-photos/1.jpg",
    "./demo-photos/2.jpg",
    "./demo-photos/3.jpg",
    "./demo-photos/4.jpg",
    "./demo-photos/5.jpg",
  ];

  const getLayoutConfig = () => {
    const urlParams = new URLSearchParams(window.location.search);
    const demoMode = urlParams.get("demo") === "1" || urlParams.get("demo") === "true";
    const demoPhotosMode =
      urlParams.get("demoPhotos") === "1" ||
      urlParams.get("demoPhotos") === "true" ||
      demoMode;

    return {
      demoMode,
      demoPhotosMode,
    };
  };

  const createState = () => ({
    projectHandle: null,
    photoHandle: null,
    projectDoc: null,
    sidecarDoc: null,
    photoRootName: "",
    tagOptions: null,
    tagFilters: { report: "", observations: "", training: "" },
    activeTagFilters: { report: [], observations: [], training: [] },
    photos: [],
    filterMode: "all",
    currentIndex: 0,
    currentPhotoUrl: "",
    currentPhotoToken: 0,
    showTagCounts: false,
  });

  const createRuntime = () => ({
    autosaveTimer: null,
    saveQueue: Promise.resolve(),
    renderTimer: null,
  });

  window.AutoBerichtPhotoSorterState = {
    IMAGE_EXTENSIONS,
    RESIZE_MAX,
    RESIZE_QUALITY,
    DEFAULT_TAGS,
    DEMO_PHOTO_URLS,
    getLayoutConfig,
    createState,
    createRuntime,
  };
})();
