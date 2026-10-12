
import axios from "axios";
import { useState } from "react";
import * as XLSX from "xlsx";

function App() {
  const [msg, setmsg] = useState("");
  const [status, setstatus] = useState(false);
  const [emailList, setEmailList] = useState([]);

  function handlemsg(evt) {
    setmsg(evt.target.value);
  }

  function handlefile(event) {
    const file = event.target.files[0];

    if (!file) return;

    const reader = new FileReader();

    reader.onload = function (e) {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array" });

        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];

        const rows = XLSX.utils.sheet_to_json(worksheet, {
          header: 1,
          defval: ""
        });

        // Read email addresses from Column A
        const validEmails = [
          ...new Set(
            rows
              .map((row) => String(row[0] || "").trim())
              .filter((email) =>
                /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)
              )
          )
        ];

        setEmailList(validEmails);

        if (validEmails.length === 0) {
          alert("No valid email addresses found in Column A.");
        } else {
          alert(`${validEmails.length} valid email address(es) loaded.`);
        }
      } catch (error) {
        console.error("Excel reading error:", error);
        alert("Unable to read this Excel file.");
        setEmailList([]);
      }
    };

    reader.onerror = function () {
      alert("Unable to read the selected file.");
      setEmailList([]);
    };

    reader.readAsArrayBuffer(file);
  }

  async function send() {
    if (!msg.trim()) {
      alert("Please enter the email message.");
      return;
    }

    if (emailList.length === 0) {
      alert("Please upload an Excel file containing email addresses.");
      return;
    }

    setstatus(true);

    try {
      const response = await axios.post(
        "https://bulkmail-backend-mgsk.onrender.com/sendemail",
        {
          message: msg.trim(),
          emailList: emailList
        }
      );

      if (response.data?.success === true) {
        alert(
          response.data.message || "Emails sent successfully!"
        );
      } else {
        alert(
          response.data?.error ||
          "The server did not confirm that the emails were sent."
        );
      }
    } catch (error) {
      console.error(
        "Email sending error:",
        error.response?.data || error.message
      );

      if (error.response) {
        alert(
          error.response.data?.error ||
          `Backend Error: ${error.response.status}`
        );
      } else {
        alert(
          "Unable to connect to the backend. Please try again."
        );
      }
    } finally {
      setstatus(false);
    }
  }

  return (
    <div>
      <div className="bg-blue-950 text-white text-center">
        <h1 className="text-2xl font-medium px-5 py-3">
          BulkMail
        </h1>
      </div>

      <div className="bg-blue-800 text-white text-center">
        <h1 className="font-medium px-5 py-3">
          We can help your business with sending multiple emails at once
        </h1>
      </div>

      <div className="bg-blue-600 text-white text-center">
        <h1 className="font-medium px-5 py-3">
          Drag and Drop
        </h1>
      </div>

      <div className="bg-blue-400 flex flex-col items-center text-black px-5 py-3">
        <textarea
          onChange={handlemsg}
          value={msg}
          className="w-[80%] h-32 py-2 outline-none px-2 border border-black rounded-md"
          placeholder="Enter the email text..."
        />

        <div>
          <input
            type="file"
            accept=".xlsx,.xls,.csv"
            onChange={handlefile}
            className="border-4 border-dashed py-4 px-4 mt-5 mb-5"
          />
        </div>

        <p>Total valid emails in the file: {emailList.length}</p>

        <button
          onClick={send}
          disabled={status}
          className="mt-2 bg-blue-950 py-2 px-4 text-white font-medium rounded-md w-fit disabled:opacity-50"
        >
          {status ? "Sending..." : "Send"}
        </button>
      </div>

      <div className="bg-blue-300 text-white text-center p-8" />
      <div className="bg-blue-200 text-white text-center p-8" />
    </div>
  );
}

export default App;
