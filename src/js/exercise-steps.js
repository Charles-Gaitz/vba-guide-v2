document.querySelectorAll('.exercise-step').forEach(function (step) {
  var header = step.querySelector('.step-header');
  var body = step.querySelector('.step-body');
  if (!header || !body) return;

  header.addEventListener('click', function () {
    var isOpen = step.classList.contains('open');
    step.classList.toggle('open');
    body.hidden = isOpen;
    header.setAttribute('aria-expanded', String(!isOpen));
  });
});

document.querySelectorAll('.exercise-hint').forEach(function (hint) {
  var btn = hint.querySelector('button');
  var body = hint.querySelector('.hint-body');
  if (!btn || !body) return;

  btn.addEventListener('click', function () {
    var isHidden = body.hidden;
    body.hidden = !isHidden;
    btn.setAttribute('aria-expanded', String(isHidden));
    btn.textContent = isHidden ? 'Hide Hint' : 'Show Hint';
  });
});

document.querySelectorAll('.exercise-solution').forEach(function (sol) {
  var btn = sol.querySelector('button');
  var body = sol.querySelector('.solution-body');
  if (!btn || !body) return;

  btn.addEventListener('click', function () {
    var isHidden = body.hidden;
    body.hidden = !isHidden;
    btn.setAttribute('aria-expanded', String(isHidden));
    btn.textContent = isHidden ? 'Hide Solution' : 'Show Solution';
  });
});
