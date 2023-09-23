const axios = require("axios");
const excel = require("exceljs");

const fetchData = async (number) => {
  try {
    const url = `https://tpsjatim.kpu.go.id/api/peta/35.${number}`;
    const headers = {
      accept: "application/json, text/javascript, */*; q=0.01",
      "accept-language": "en-US,en;q=0.9,id;q=0.8",
      "sec-ch-ua":
        '"Chromium";v="116", "Not)A;Brand";v="24", "Google Chrome";v="116"',
      "sec-ch-ua-mobile": "?0",
      "sec-ch-ua-platform": '"macOS"',
      "sec-fetch-dest": "empty",
      "sec-fetch-mode": "cors",
      "sec-fetch-site": "same-origin",
      "user-agent":
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/116.0.0.0 Safari/537.36",
      "x-requested-with": "XMLHttpRequest",
      cookie: "ci_session=nnl94pglgcr2i5dmv7benplttk0f9pk4",
      referer: `https://tpsjatim.kpu.go.id/peta/35.${number}`,
    };
    console.log(
      `Mengambil data dari https://tpsjatim.kpu.go.id/peta/35.${number}`
    );
    let { data } = await axios.get(url, { headers });
    response = data.data;
    console.log(
      `Berhasil fetch data dari ${number}, sampel data pertama ->`,
      response[0]
    );
    return response;
  } catch (error) {
    console.log(`Gagal fetch data dari ${number}, data dilewati..`, error);
  }
};

const fetchAllData = async () => {
  const result = [];

  for (let i = 1; i < 80; i++) {
    const number = i.toString().padStart(2, "0");
    const data = await fetchData(number);
    if (data) {
      result.push(data);
    }
  }

  return result;
};

fetchAllData()
  .then((result) => {
    const workbook = new excel.Workbook();
    const worksheet = workbook.addWorksheet("Data tpsjatim TPU");

    worksheet.columns = [
      { header: "Kota / Kab", key: "kotakab", width: 10 },
      { header: "Kecamatan", key: "kecamatan", width: 10 },
      { header: "Desa / Kel.", key: "desakel", width: 10 },
      { header: "Kode Desa / Kel", key: "kode_desakel", width: 10 },
      { header: "Alamat", key: "alamat", width: 10 },
      { header: "No. TPS", key: "no_tps", width: 10 },
      { header: "latitude", key: "latitude", width: 10 },
      { header: "longitude", key: "longitude", width: 10 },
      { header: "Pemilih Laki-laki", key: "jumlah_pria", width: 10 },
      { header: "Pemilih Perempuan", key: "jumlah_wanita", width: 10 },
      { header: "Total Pemilih", key: "jumlah_pemilih", width: 10 },
    ];

    // Concatenate response.data into one array
    const mergedData = [].concat(...result);

    worksheet.addRows(mergedData);

    workbook.xlsx
      .writeFile("Data tpsjatim TPU.xlsx")
      .then(() => {
        console.log("Berhasil membuat file excel");
      })
      .catch((error) => {
        console.error(error);
      });
  })
  .catch((error) => {
    console.error(error);
  });
