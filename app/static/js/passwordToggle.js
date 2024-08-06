const showPsswdtoggle = (id, icon) => {
  const input = document.getElementById(id);
  if (input.type === 'password') {
      input.type = 'text';
      icon.className = icon.className.replace('bi bi-eye', 'bi bi-eye-slash');
  } else {
      input.type = 'password';
      icon.className = icon.className.replace('bi bi-eye-slash', 'bi bi-eye');
  }
};


