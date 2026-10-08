const avatar = document.getElementById("avatarBtn");
      const dropdown = document.getElementById("userDropdown");
      avatar.addEventListener("click", () => dropdown.classList.toggle("show"));
      document.addEventListener("click", (e) => {
        if (!e.target.closest(".user-profile")) dropdown.classList.remove("show");
      });
      document.querySelectorAll(".menu .nav-link").forEach((link) => {
        link.addEventListener("click", (e) => {
          e.preventDefault();
          document.querySelectorAll(".menu .nav-link").forEach((l) => l.classList.remove("active"));
          link.classList.add("active");
        });
      });