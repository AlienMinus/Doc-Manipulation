document.addEventListener("DOMContentLoaded", function () {
  const loadingOverlay = document.getElementById("loading-overlay");
  const resultsSection = document.getElementById("results-section");
  const resultsBody = document.getElementById("results-body");
  const clearBtn = document.getElementById("clear-results");
  const copyBtn = document.getElementById("copy-results");

  clearBtn.addEventListener("click", () => {
    resultsSection.style.display = "none";
    resultsBody.innerHTML = "";
  });

  copyBtn.addEventListener("click", () => {
    let textToCopy = "";
    const textarea = resultsBody.querySelector("textarea");
    const codeBlock = resultsBody.querySelector("code");

    if (textarea) {
      textToCopy = textarea.value;
    } else if (codeBlock) {
      textToCopy = codeBlock.textContent;
    } else {
      textToCopy = resultsBody.innerText;
    }

    if (textToCopy) {
      navigator.clipboard
        .writeText(textToCopy)
        .then(() => {
          const originalHtml = copyBtn.innerHTML;
          copyBtn.innerHTML = '<i class="bi bi-check-lg"></i>';
          copyBtn.classList.remove("btn-outline-secondary");
          copyBtn.classList.add("btn-success");

          setTimeout(() => {
            copyBtn.innerHTML = originalHtml;
            copyBtn.classList.remove("btn-success");
            copyBtn.classList.add("btn-outline-secondary");
          }, 2000);
        })
        .catch((err) => {
          console.error("Failed to copy: ", err);
          alert("Failed to copy to clipboard");
        });
    }
  });

  const forms = document.querySelectorAll("form");
  forms.forEach((form) => {
    // Set API endpoints based on the hidden input 'feature'
    const featureInput = form.querySelector('input[name="feature"]');
    if (featureInput) {
      form.setAttribute("data-endpoint", "/api/" + featureInput.value);
    }

    form.addEventListener("submit", async (e) => {
      e.preventDefault();
      loadingOverlay.style.display = "flex";
      resultsSection.style.display = "none";
      resultsBody.innerHTML = "";

      const formData = new FormData(form);

      // Capture the submitter button value (for Preview vs Download)
      if (e.submitter && e.submitter.name) {
        formData.append(e.submitter.name, e.submitter.value);
      }

      const endpoint = form.getAttribute("data-endpoint");

      try {
        const response = await fetch(endpoint, {
          method: "POST",
          body: formData,
        });

        const contentType = response.headers.get("content-type");

        if (!response.ok) {
          let errorMsg = "An error occurred";
          if (contentType && contentType.includes("application/json")) {
            const errData = await response.json();
            errorMsg = errData.error || errorMsg;
          } else {
            errorMsg = `Error ${response.status}: ${response.statusText}`;
          }
          throw new Error(errorMsg);
        }

        if (contentType.includes("application/json")) {
          const data = await response.json();
          renderJSONResults(data);
          resultsSection.style.display = "block";
        } else {
          // Assume blob/file download
          const blob = await response.blob();
          const url = window.URL.createObjectURL(blob);
          const a = document.createElement("a");
          a.href = url;
          // Try to get filename from header or default
          const disposition = response.headers.get("content-disposition");
          let filename = "download.docx";
          if (disposition && disposition.indexOf("attachment") !== -1) {
            const filenameRegex = /filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/;
            const matches = filenameRegex.exec(disposition);
            if (matches != null && matches[1]) {
              filename = matches[1].replace(/['"]/g, "");
            }
          }
          a.download = filename;
          document.body.appendChild(a);
          a.click();
          window.URL.revokeObjectURL(url);
          a.remove();
        }
      } catch (error) {
        alert(error.message);
      } finally {
        loadingOverlay.style.display = "none";
      }
    });
  });

  function renderJSONResults(data) {
    if (data.metadata) {
      let html = '<table class="table table-striped"><tbody>';
      for (const [key, value] of Object.entries(data.metadata)) {
        html += `<tr><th scope="row" style="width: 30%">${key}</th><td>${value}</td></tr>`;
      }
      html += "</tbody></table>";
      resultsBody.innerHTML = html;
    } else if (data.text) {
      resultsBody.innerHTML = `<textarea class="form-control" rows="10">${data.text}</textarea>`;
    } else if (typeof data.markdown !== "undefined") {
      // Escape HTML to ensure it displays as code
      const escaped = (data.markdown || "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;");
      resultsBody.innerHTML = `<pre><code class="language-markdown" style="border-radius: 8px;">${escaped}</code></pre>`;
      if (window.hljs) {
        hljs.highlightElement(resultsBody.querySelector("code"));
      }
    } else if (data.images) {
      let html = '<div class="row">';
      data.images.forEach((img) => {
        html += `
                    <div class="col-md-4 mb-4">
                        <div class="card h-100">
                            <img src="data:${img.mime};base64,${img.data}" class="card-img-top" style="max-height: 200px; object-fit: contain; padding: 10px;">
                            <div class="card-body text-center">
                                <p class="card-text small text-muted">${img.filename}</p>
                            </div>
                        </div>
                    </div>`;
      });
      html += "</div>";
      resultsBody.innerHTML = html;
    } else if (data.tables) {
      let html = "";
      data.tables.forEach((table, index) => {
        html += `<h5>Table ${index + 1}</h5><table class="table table-bordered table-sm mb-4">`;
        table.forEach((row) => {
          html += "<tr>";
          row.forEach((cell) => {
            html += `<td>${cell}</td>`;
          });
          html += "</tr>";
        });
        html += "</table>";
      });
      resultsBody.innerHTML = html;
    }
  }
});
