const dialog = document.querySelector("#image-dialog");
const dialogImage = document.querySelector("#dialog-image");
const dialogCaption = document.querySelector("#dialog-caption");

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

document.querySelector(".dialog-close").addEventListener("click", () => dialog.close());
dialog.addEventListener("click", (event) => {
  if (event.target === dialog) dialog.close();
});
