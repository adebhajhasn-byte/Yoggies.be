// Swap main image when a thumbnail is clicked
  function phSwap(thumb) {
    document.querySelectorAll('.ph-thumb').forEach(function(t) {
      t.classList.remove('active');
    });
    thumb.classList.add('active');
    var thumbImg = thumb.querySelector('img');
    var mainImg  = document.querySelector('#phMainImg img');
    if (thumbImg && mainImg) {
      // Thumbs are 200px; data-full is the 600px rendition. Drop the hero's
      // srcset first, or the browser keeps picking from the old candidate list
      // and ignores the new src entirely.
      mainImg.removeAttribute('srcset');
      mainImg.removeAttribute('sizes');
      mainImg.src = thumbImg.dataset.full || thumbImg.src;
      mainImg.alt = thumbImg.alt;
    }
  }