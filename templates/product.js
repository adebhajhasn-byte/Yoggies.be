// Swap main image when a thumbnail is clicked
  function phSwap(thumb) {
    document.querySelectorAll('.ph-thumb').forEach(function(t) {
      t.classList.remove('active');
    });
    thumb.classList.add('active');
    var thumbImg = thumb.querySelector('img');
    var mainImg  = document.querySelector('#phMainImg img');
    if (thumbImg && mainImg) {
      mainImg.src = thumbImg.src;
      mainImg.alt = thumbImg.alt;
    }
  }