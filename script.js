document.addEventListener("DOMContentLoaded", () => {
    fetchProducts();
});

async function fetchProducts() {
    try {
        const searchQuery = document.getElementById("search").value.trim();
        const searchMessage = document.getElementById("searchMessage");
        let url = "/products";

        if (searchQuery) {
            url += `?search=${encodeURIComponent(searchQuery)}`;
        }

        const res = await fetch(url);
        const products = await res.json();
        console.log("Отримані товари:", products);

        renderProducts(products);

        // Якщо є результати пошуку, скролимо до таблиці
        if (searchQuery && products.length > 0) {
            const table = document.querySelector("table");
            table.scrollIntoView({ behavior: "smooth" });
        }

        // Повідомлення про відсутність товару
        if (searchQuery && products.length === 0) {
            searchMessage.textContent = "Такого товару не знайдено";
        } else {
            searchMessage.textContent = ""; // прибираємо повідомлення
        }

    } catch (error) {
        console.error("Помилка отримання товарів:", error);
    }
}


function renderProducts(products) {
    const tableBody = document.getElementById("productTableBody");
    tableBody.innerHTML = "";
    products.forEach(addProductToTable);
}

function addProductToTable(product) {
    const tableBody = document.getElementById("productTableBody");

    const row = document.createElement("tr");
    row.setAttribute("data-id", product.id); // Додаємо ID товару

    row.innerHTML = `
        <td class="name">${product.name}</td>
        <td>${product.image ? `<img src="${product.image}" width="50">` : "—"}</td>
        <td class="material">${product.material || "—"}</td>
        <td class="size">${product.size || "—"}</td>
        <td class="description">${product.description || "—"}</td>
        <td class="manufacturer">${product.manufacturer || "—"}</td>
        <td class="quantity">${product.quantity}</td>
        <td class="price">${Number(product.price).toFixed(2)} грн</td>
        <td class="total_price">${Number(product.total_price).toFixed(2)} грн</td>
        <td class="delivery_date">${product.delivery_date || "—"}</td>
        <td class="supplier">${product.supplier || "—"}</td>
        <td>${product.availability}</td>
        <td>
            <button onclick="editProduct(${product.id})">Редагувати</button>
            <button onclick="deleteProduct(${product.id})">Видалити</button>
        </td>
    `;

    tableBody.appendChild(row);
}

async function addProduct() {
    const form = document.getElementById("productForm");
    const formData = new FormData(form);

    if (!formData.get("name") || !formData.get("quantity") || !formData.get("price")) {
        alert("Назва, кількість і ціна є обов'язковими полями!");
        return;
    }

    try {
        const res = await fetch("/products", { 
            method: "POST", 
            body: formData 
        });

        const product = await res.json();
        if (!res.ok) throw new Error(product.error || "Помилка додавання товару");

        console.log("Товар додано:", product);
        fetchProducts(); // Оновлення списку товарів
        form.reset(); // Очистка форми після додавання
    } catch (error) {
        console.error("Помилка додавання товару:", error);
        alert("Не вдалося додати товар! Перевір консоль.");
    }
}

async function deleteProduct(id) {
    if (!confirm("Ви впевнені, що хочете видалити цей товар?")) return;

    try {
        const res = await fetch(`/products/${id}`, { method: "DELETE" });
        const result = await res.json();

        if (result.deleted) {
            alert("Товар успішно видалено!");
            fetchProducts(); // Оновлення таблиці після видалення
        } else {
            alert("Помилка видалення товару!");
        }
    } catch (error) {
        console.error("Помилка видалення:", error);
    }
}

function editProduct(id) {
    const row = document.querySelector(`tr[data-id='${id}']`);
    if (!row) return;

    document.getElementById("productId").value = id;
    document.getElementById("name").value = row.querySelector(".name").innerText;
    document.getElementById("material").value = row.querySelector(".material").innerText;
    document.getElementById("size").value = row.querySelector(".size").innerText;
    document.getElementById("description").value = row.querySelector(".description").innerText;
    document.getElementById("manufacturer").value = row.querySelector(".manufacturer").innerText;
    document.getElementById("quantity").value = row.querySelector(".quantity").innerText;
    document.getElementById("price").value = row.querySelector(".price").innerText.replace(" грн", "");
    document.getElementById("delivery_date").value = row.querySelector(".delivery_date").innerText;
    document.getElementById("supplier").value = row.querySelector(".supplier").innerText;

    window.scrollTo({ top: 0, behavior: "smooth" });
}

async function saveProduct(event) {
    event.preventDefault();
    const form = document.getElementById("productForm");
    const formData = new FormData(form);
    const productId = document.getElementById("productId").value;
    const url = productId ? `/products/${productId}` : "/products";
    const method = productId ? "PUT" : "POST";

    try {
        const res = await fetch(url, { 
            method, 
            body: formData 
        });

        const product = await res.json();
        if (!res.ok) throw new Error(product.error || "Помилка збереження товару");

        console.log("Товар успішно збережено:", product);
        fetchProducts();
        form.reset();
    } catch (error) {
        console.error("Помилка додавання товару:", error);
        alert("Не вдалося додати товар! Перевір консоль.");
    }
}
document.getElementById("productForm").addEventListener("submit", saveProduct)

async function filterProducts() {
    const minPrice = document.getElementById("minPrice").value;
    const maxPrice = document.getElementById("maxPrice").value;
    const manufacturer = document.getElementById("manufacturerFilter").value;
    const availability = document.getElementById("availabilityFilter").value;

    let url = `/products?`;

    if (minPrice) url += `minPrice=${minPrice}&`;
    if (maxPrice) url += `maxPrice=${maxPrice}&`;
    if (manufacturer) url += `manufacturer=${encodeURIComponent(manufacturer)}&`;
    if (availability) url += `availability=${encodeURIComponent(availability)}&`;

    try {
        const res = await fetch(url);
        const products = await res.json();
        renderProducts(products);
    } catch (error) {
        console.error("Помилка фільтрації:", error);
    }
}

function resetFilters() {
    document.getElementById("minPrice").value = "";
    document.getElementById("maxPrice").value = "";
    document.getElementById("manufacturerFilter").value = "";
    document.getElementById("availabilityFilter").value = "";
    fetchProducts(); // Завантажує всі товари без фільтрів
}

async function loadManufacturers() {
    try {
        const res = await fetch("/manufacturers");
        const manufacturers = await res.json();
        const select = document.getElementById("manufacturerFilter");

        manufacturers.forEach(manufacturer => {
            const option = document.createElement("option");
            option.value = manufacturer;
            option.textContent = manufacturer;
            select.appendChild(option);
        });
    } catch (error) {
        console.error("Помилка завантаження виробників:", error);
    }
}
document.addEventListener("DOMContentLoaded", loadManufacturers);

async function fetchSupplyData() {
    const startDate = document.getElementById("startDate").value;
    const endDate = document.getElementById("endDate").value;
    
    if (!startDate || !endDate) {
        alert("Оберіть період для побудови графіка!");
        return;
    }

    try {
        const res = await fetch(`/supply-data?startDate=${startDate}&endDate=${endDate}`);
        const data = await res.json();

        if (!res.ok) throw new Error(data.error || "Помилка отримання даних");

        console.log("Дані для графіка:", data);
        renderSupplyChart(data);
    } catch (error) {
        console.error("Помилка завантаження графіка:", error);
    }
}
let supplyChart;

function renderSupplyChart(data) {
    const ctx = document.getElementById("supplyChart").getContext("2d");
    
    if (supplyChart) supplyChart.destroy();

    supplyChart = new Chart(ctx, {
        type: "bar",
        data: {
            labels: data.map(item => item.delivery_date),
            datasets: [{
                label: "Надходження товарів",
                data: data.map(item => item.total_quantity),
                backgroundColor: "rgba(54, 162, 235, 0.5)",
                borderColor: "rgba(54, 162, 235, 1)",
                borderWidth: 1
            }]
        },
        options: {
            responsive: true,
            scales: {
                x: { title: { display: true, text: "Дата поставки" } },
                y: { title: { display: true, text: "Кількість товарів" }, beginAtZero: true }
            }
        }
    });
}

// === ЕКСПОРТ EXCEL ===
function exportExcel() {
    window.location.href = "/export/excel";
}

// === ІМПОРТ EXCEL ===
async function importExcel(event) {
    const file = event.target.files[0];
    if (!file) return alert("Оберіть файл Excel для імпорту!");

    const formData = new FormData();
    formData.append("file", file);

    const res = await fetch("/import/excel", { method: "POST", body: formData });
    const result = await res.json();

    if (res.ok) {
        alert(`Імпортовано ${result.imported} товарів`);
        fetchProducts();
    } else {
        alert("Помилка імпорту: " + (result.error || "невідома помилка"));
    }
}

function exportTxt() {
    window.location.href = "/export/txt";
}

