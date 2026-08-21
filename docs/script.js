const isEnglish = document.documentElement.lang === "en";
const assetPrefix = isEnglish ? "../" : "";
const pageCopy = isEnglish
  ? [
      {
        title: "Page 1 — source document",
        beforeDescription: "The source image is placed on a regular white page.",
        afterDescription: "The page is now black and white and includes a binding element.",
      },
      {
        title: "Page 2 — electronic page",
        beforeDescription: "A clean electronic page created in Word.",
        afterDescription: "A paper fold and ribbon make the page look like a scan.",
      },
      {
        title: "Page 3 — final part",
        beforeDescription: "A separately scanned final page of the document.",
        afterDescription: "The final page is included in the finished PDF with a consistent appearance.",
      },
    ]
  : [
      {
        title: "Страница 1 — исходный документ",
        beforeDescription: "Изображение находится на обычной белой странице.",
        afterDescription: "Страница стала чёрно-белой, а сверху появился элемент прошивки.",
      },
      {
        title: "Страница 2 — электронная страница",
        beforeDescription: "Чистая электронная страница, созданная в Word.",
        afterDescription: "Добавлены сгиб листа и лента, изображение стало похоже на скан.",
      },
      {
        title: "Страница 3 — последняя часть",
        beforeDescription: "Отдельно отсканированная последняя страница документа.",
        afterDescription: "Последняя страница приведена к общему виду и добавлена в итоговый PDF.",
      },
    ];

const pageFiles = [
  {
    before: "assets/До/Перевод паспорта гражданина Листенбурга_pymu_page_001.png",
    after: "assets/После/Перевод паспорта гражданина Листенбурга_page_001.png",
  },
  {
    before: "assets/До/Перевод паспорта гражданина Листенбурга_pymu_page_002.png",
    after: "assets/После/Перевод паспорта гражданина Листенбурга_page_002.png",
  },
  {
    before: "assets/До/not-notary-translation-light.png",
    after: "assets/После/Перевод паспорта гражданина Листенбурга_page_003.png",
  },
];

const pages = pageCopy.map((copy, index) => ({
  ...copy,
  before: assetPrefix + pageFiles[index].before,
  after: assetPrefix + pageFiles[index].after,
}));

let currentPage = 0;
let currentView = "before";

const comparisonImage = document.querySelector("#comparison-image");
const pageTitle = document.querySelector("#page-title");
const pageDescription = document.querySelector("#page-description");
const pageCounter = document.querySelector("#page-counter");
const switchButtons = [...document.querySelectorAll(".switch-button")];
const pageDots = [...document.querySelectorAll(".page-dot")];
const dialog = document.querySelector("#image-dialog");
const dialogImage = document.querySelector("#dialog-image");
const dialogCaption = document.querySelector("#dialog-caption");

function renderPage() {
  const page = pages[currentPage];
  const isBefore = currentView === "before";

  comparisonImage.src = page[currentView];
  comparisonImage.alt = isEnglish
    ? `${page.title}, ${isBefore ? "before" : "after"} processing`
    : `${page.title}, ${isBefore ? "до" : "после"} обработки`;
  pageTitle.textContent = page.title;
  pageDescription.textContent = isBefore ? page.beforeDescription : page.afterDescription;
  pageCounter.textContent = `${currentPage + 1} / ${pages.length}`;

  switchButtons.forEach((button) => {
    const isActive = button.dataset.view === currentView;
    button.classList.toggle("is-active", isActive);
    button.setAttribute("aria-pressed", String(isActive));
  });

  pageDots.forEach((dot, index) => {
    const isActive = index === currentPage;
    dot.classList.toggle("is-active", isActive);
    if (isActive) {
      dot.setAttribute("aria-current", "true");
    } else {
      dot.removeAttribute("aria-current");
    }
  });
}

function changePage(delta) {
  currentPage = (currentPage + delta + pages.length) % pages.length;
  renderPage();
}

function openCurrentImage() {
  const page = pages[currentPage];
  dialogImage.src = page[currentView];
  dialogImage.alt = comparisonImage.alt;
  dialogCaption.textContent = isEnglish
    ? `${page.title} — ${currentView === "before" ? "before" : "after"} processing`
    : `${page.title} — ${currentView === "before" ? "до обработки" : "после обработки"}`;
  dialog.showModal();
}

switchButtons.forEach((button) => {
  button.addEventListener("click", () => {
    currentView = button.dataset.view;
    renderPage();
  });
});

pageDots.forEach((dot) => {
  dot.addEventListener("click", () => {
    currentPage = Number(dot.dataset.page);
    renderPage();
  });
});

document.querySelector(".page-arrow.previous").addEventListener("click", () => changePage(-1));
document.querySelector(".page-arrow.next").addEventListener("click", () => changePage(1));
document.querySelector(".page-image-button").addEventListener("click", openCurrentImage);

document.querySelectorAll("[data-lightbox]").forEach((button) => {
  button.addEventListener("click", () => {
    const image = button.querySelector("img");
    dialogImage.src = image.src;
    dialogImage.alt = image.alt;
    dialogCaption.textContent = button.dataset.caption || image.alt;
    dialog.showModal();
  });
});

document.querySelectorAll("[data-carousel]").forEach((carousel) => {
  const slides = [...carousel.querySelectorAll(".carousel-slide")];
  const counter = carousel.querySelector("[data-carousel-counter]");
  let index = 0;

  const showSlide = (nextIndex) => {
    index = (nextIndex + slides.length) % slides.length;
    slides.forEach((slide, slideIndex) => slide.classList.toggle("is-active", slideIndex === index));
    counter.textContent = `${index + 1} / ${slides.length}`;
  };

  carousel.querySelector("[data-carousel-prev]").addEventListener("click", () => showSlide(index - 1));
  carousel.querySelector("[data-carousel-next]").addEventListener("click", () => showSlide(index + 1));
});

document.querySelectorAll("[data-copy]").forEach((button) => {
  button.addEventListener("click", async () => {
    try {
      await navigator.clipboard.writeText(button.dataset.copy);
      const icon = button.querySelector("span");
      icon.textContent = "✓";
      button.title = button.dataset.copiedLabel;
      button.setAttribute("aria-label", button.dataset.copiedLabel);
      window.setTimeout(() => {
        icon.textContent = "⧉";
        button.title = button.dataset.copyLabel;
        button.setAttribute("aria-label", button.dataset.copyLabel);
      }, 1600);
    } catch {
      window.prompt(button.dataset.copyLabel, button.dataset.copy);
    }
  });
});

document.querySelectorAll("[data-open-page]").forEach((stack) => {
  stack.addEventListener("click", () => {
    currentView = stack.dataset.openPage;
    currentPage = Number(stack.dataset.pageIndex || 0);
    renderPage();
    document.querySelector("#how-it-works").scrollIntoView({ behavior: "smooth" });
  });
});

document.querySelector(".dialog-close").addEventListener("click", () => dialog.close());
dialog.addEventListener("click", (event) => {
  if (event.target === dialog) dialog.close();
});

document.addEventListener("keydown", (event) => {
  if (dialog.open) return;
  if (event.key === "ArrowLeft") changePage(-1);
  if (event.key === "ArrowRight") changePage(1);
});

renderPage();
