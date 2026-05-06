document.querySelectorAll('.mc-question').forEach(function (question) {
  var options = question.querySelectorAll('.mc-option');
  var explanation = question.querySelector('.mc-explanation');

  options.forEach(function (option) {
    option.addEventListener('click', function () {
      if (this.disabled) return;

      var correct = question.querySelector('.mc-option[data-correct="true"]');

      if (correct) correct.classList.add('correct');
      if (this !== correct) this.classList.add('incorrect');

      options.forEach(function (opt) {
        opt.disabled = true;
        opt.setAttribute('aria-disabled', 'true');
      });

      if (explanation) explanation.hidden = false;
    });
  });
});
