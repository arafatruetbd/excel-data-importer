

# 📊 Excel Plugin with Office JavaScript API

A simple yet functional Excel Add-in that demonstrates how to fetch and display external data in an Excel worksheet using the [Office JavaScript API](https://learn.microsoft.com/office/dev/add-ins/overview/office-add-ins). This plugin is ideal for learning the basics of Excel Add-in development and Office integration using JavaScript.

## ✨ Features

* ✅ Fetches sample data from a mock API
* ✅ Automatically inserts the data into the current Excel worksheet
* ✅ Built using the Office JS API and Fluent UI
* ✅ Clean, user-friendly interface
* ✅ Ready for further extension and real-world integrations

## 📸 Preview

![Plugin Screenshot](https://github.com/arafatruetbd/excel-data-importer/blob/master/ExcelDataImporter/assets/Screenshot%202025-05-25%20195905.png)
*A welcome page with a "Run" button to fetch and populate data into Excel*

## 🚀 Getting Started

### 1. Prerequisites

* Node.js (>=14.x)
* [Office Add-in development environment](https://learn.microsoft.com/office/dev/add-ins/overview/office-add-ins)
* Excel (Microsoft 365) Desktop or Web version

### 2. Install Office Add-in CLI tools

```bash
yarn global add yo generator-office
```

### 3. Clone the repository

```bash
git clone https://github.com/your-username/excel-plugin-office-js.git
cd excel-plugin-office-js
```

### 4. Run the Add-in locally

```bash
yarn 
yarn start
```

This will sideload the add-in in Excel and open it in the task pane.

> 💡 If sideloading doesn’t work, follow [this guide](https://learn.microsoft.com/office/dev/add-ins/testing/test-debug-office-add-ins#sideload-an-office-add-in-for-testing) to manually sideload your add-in.

## 🧪 How It Works

The add-in:

1. Displays a welcome interface with a "Run" button.
2. On click, triggers a mock API call to simulate fetching external data.
3. Inserts the returned data directly into the active worksheet in Excel.

## 🧰 Tech Stack

* JavaScript
* Office JS API
* Fluent UI
* Excel Desktop/Web


## 📄 License

This project is licensed under the [MIT License](./LICENSE).

---

## 🤝 Contributing

Pull requests are welcome! Feel free to fork and improve this example — especially if you'd like to integrate real APIs, improve UI/UX, or add error handling.

---

## 👨‍💻 Author

**Arfat Hossain**
📫 [LinkedIn](https://www.linkedin.com/in/arfat-hossain-a89531148) | 🌐 [Portfolio](https://portfolio-arafat.vercel.app/)

