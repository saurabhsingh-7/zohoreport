const axios = require('axios');
const path = require('path')
const os = require('os');
const https = require('https');
const homeDir = os.homedir();
const xlsx = require('xlsx');
const ExcelJS = require('exceljs');
const zohoClientId = '1000.FM3AZOWDHHSWVIBS25Z0OKG3HMJP5C';
const zohoClientSecret = 'd872a47ddec2f7643fba24b57944ca7b732d68ab77';
// const zohoRefreshToken = '1000.765ad5893706d79df4ea22d3e1e55af7.514d3815dec9359d5dacd4ce85a342b0';
const zohoRefreshToken = '1000.a2e8f7263653f17d9684dbfd4549f45c.2c929542e520239812269fcbfa4acac7';
const sheetId = 'u0o741d8da6feb87e4f35bee1769da9d10837';
const sheetName = 'Sheet1';
async function getZohoAccessToken() {
  const url = 'https://accounts.zoho.com/oauth/v2/token';
  const data = {
    grant_type: 'refresh_token',
    client_id: zohoClientId,
    client_secret: zohoClientSecret,
    refresh_token: zohoRefreshToken,
  };
  const headers = { 'Content-Type': 'application/x-www-form-urlencoded' };
  try {
    const response = await axios.post(url, new URLSearchParams(data).toString(), { headers });
    return response.data.access_token;
  } catch (error) {
    console.error('Error fetching Zoho access token:', error);
    throw error;
  }
}
async function fetchSalesOrders(accessToken) {
  const config = {
    method: 'get',
    maxBodyLength: Infinity,
    url: 'https://www.zohoapis.com/books/v3/salesorders?organization_id=773203695',
    headers: {
      'Authorization': `Zoho-oauthtoken ${accessToken}`,
      'Content-Type': 'application/json',
    }
  };

  try {
    const response = await axios.request(config);
    return response.data.salesorders;
  } catch (error) {
    console.error('Error fetching sales orders:', error);
    throw error;
  }
}
async function fetchPurchaseOrders(accessToken) {
  const config = {
    method: 'get',
    maxBodyLength: Infinity,
    url: 'https://www.zohoapis.com/books/v3/purchaseorders?organization_id=773203695',
    headers: {
      'Authorization': `Zoho-oauthtoken ${accessToken}`,
      'Content-Type': 'application/json',
    }
  };
  try {
    const response = await axios.request(config);
    return response.data.purchaseorders;
  } catch (error) {
    console.error('Error fetching purchase orders:', error);
    throw error;
  }
}

async function fetchinvoices(accessToken) {
  const config = {
    method: 'get',
    maxBodyLength: Infinity,
    url: 'https://www.zohoapis.com/books/v3/invoices?organization_id=773203695',
    headers: {
      'Authorization': `Zoho-oauthtoken ${accessToken}`,
      'Content-Type': 'application/json',
    }
  };
  try {
    const response = await axios.request(config);
    // console.log(response.data.invoices);
    return response.data.invoices;
  } catch (error) {
    console.error('Error fetching purchase orders:', error);
    throw error;
  }
}


async function fetchbills(accessToken) {
  const config = {
    method: 'get',
    maxBodyLength: Infinity,
    url: 'https://www.zohoapis.com/books/v3/bills?organization_id=773203695',
    headers: {
      'Authorization': `Zoho-oauthtoken ${accessToken}`,
      'Content-Type': 'application/json',
    }
  };
  try {
    const response = await axios.request(config);
    console.log(response.data.bills);
    return response.data.bills;
  } catch (error) {
    console.error('Error fetching purchase orders:', error);
    throw error;
  }
}






async function insertDataIntoSheet(accessToken, resultArray) {
  console.log(homeDir,"homeDir");
  const workbooks2 = new ExcelJS.Workbook();
  const worksheet2 = workbooks2.addWorksheet('Sheet 1');
  const headers2 = [
    'Customer Name',
    'SO Date',
    'SO Number',
    'Customer PO',
    'Amount',
    'Supplier PO No',
    'Supplier Name',
    'Supplier PO Value',
    'Invoice Number',
    'Invoice Value',
    'Bill Number',
    'Bill Value'
  ];
  worksheet2.addRow(headers2);
  const rows = resultArray.map(map => {
    return {
      "Customer Name": map.get("Customer Name"),
      "SO Date": map.get("SO Date"),
      "SO Number": map.get("SO Number"),
      "Customer PO": map.get("Customer PO"),
      "Amount": map.get("Amount"),
      "Supplier PO No": map.get("Supplier PO No"),
      "Supplier Name": map.get("Supplier Name"),
      "Supplier PO Value": map.get("Supplier PO Value"),
      "Invoice Number": map.get("Invoice Number"),
      "Invoice Value": map.get("Invoice Value"),
      "Bill Number": map.get("Bill Number"),
      "Bill Value": map.get("Bill Value")
    };
  });
  rows.forEach(async (row) => {
    worksheet2.addRow(Object.values(row));
  })
  const csvFilePaths2 = path.join(homeDir, 'BooksReport6.csv');
workbooks2.xlsx.writeFile(csvFilePaths2)
    .then(() => {
        console.log('Excel file created successfully csv.');
    })
    .catch(error => {
        console.error('Error creating Excel file:', error);
    });
}


async function main() {
  try {
    const accessToken = await getZohoAccessToken();

    const [salesOrders, purchaseOrders, invoicesAll, bills] = await Promise.all([
      fetchSalesOrders(accessToken),
      fetchPurchaseOrders(accessToken),
      fetchinvoices(accessToken),
      fetchbills(accessToken)
    ]);
     
    const resultArray = [];

    salesOrders.forEach(salesOrder => {
      
      const relevantPurchaseOrders = purchaseOrders.filter(purchaseOrder => purchaseOrder.reference_number === salesOrder.salesorder_number);
      const relevantInvoices = invoicesAll.filter(invoice => invoice.reference_number === salesOrder.salesorder_number);
      relevantPurchaseOrders.forEach(purchaseOrder => {
        const relevantBills = bills.filter(bill => bill.reference_number === purchaseOrder.purchaseorder_number);
        
        relevantInvoices.forEach(invoice => {
          relevantBills.forEach(bill => {
            const combinedData = new Map();
            combinedData.set("Customer Name", salesOrder.customer_name);
            combinedData.set("SO Date", salesOrder.date);
            combinedData.set("SO Number", salesOrder.salesorder_number);
            combinedData.set("Customer PO", salesOrder.reference_number);
            combinedData.set("Amount", salesOrder.total);
            combinedData.set("Supplier PO No", purchaseOrder.purchaseorder_number);
            combinedData.set("Supplier Name", purchaseOrder.vendor_name);
            combinedData.set("Invoice Number", invoice.invoice_number);
            combinedData.set("Invoice Value", invoice.total);
            combinedData.set("Bill Number", bill.bill_number);
            combinedData.set("Bill Value", bill.total);
            resultArray.push(combinedData);
          });

          if (relevantBills.length === 0) {
            const combinedData = new Map();
            combinedData.set("Customer Name", salesOrder.customer_name);
            combinedData.set("SO Date", salesOrder.date);
            combinedData.set("SO Number", salesOrder.salesorder_number);
            combinedData.set("Customer PO", salesOrder.reference_number);
            combinedData.set("Amount", salesOrder.total);
            combinedData.set("Supplier PO No", purchaseOrder.purchaseorder_number);
            combinedData.set("Supplier Name", purchaseOrder.vendor_name);
            combinedData.set("Invoice Number", invoice.invoice_number);
            combinedData.set("Invoice Value", invoice.total);
            combinedData.set("Bill Number", " ");
            resultArray.push(combinedData);
          }
        });

        if (relevantInvoices.length === 0) {
          relevantBills.forEach(bill => {
            const combinedData = new Map();
            combinedData.set("Customer Name", salesOrder.customer_name);
            combinedData.set("SO Date", salesOrder.date);
            combinedData.set("SO Number", salesOrder.salesorder_number);
            combinedData.set("Customer PO", salesOrder.reference_number);
            combinedData.set("Amount", salesOrder.total);
            combinedData.set("Supplier PO No", purchaseOrder.purchaseorder_number);
            combinedData.set("Supplier Name", purchaseOrder.vendor_name);
            combinedData.set("Invoice Number", " ");
            combinedData.set("Invoice Value", " ");
            combinedData.set("Bill Number", bill.bill_number);
            combinedData.set("Bill Value", bill.total);
            resultArray.push(combinedData);
          });

          if (relevantBills.length === 0) {
            const combinedData = new Map();
            combinedData.set("Customer Name", salesOrder.customer_name);
            combinedData.set("SO Date", salesOrder.date);
            combinedData.set("SO Number", salesOrder.salesorder_number);
            combinedData.set("Customer PO", salesOrder.reference_number);
            combinedData.set("Amount", salesOrder.total);
            combinedData.set("Supplier PO No", purchaseOrder.purchaseorder_number);
            combinedData.set("Supplier Name", purchaseOrder.vendor_name);
            combinedData.set("Invoice Number", " ");
            combinedData.set("Invoice Value", " ");
            combinedData.set("Bill Number", " ");
            combinedData.set("Bill Value", " ");
            resultArray.push(combinedData);
          }
        }
      });

      if (relevantPurchaseOrders.length === 0) {
        relevantInvoices.forEach(invoice => {
          const combinedData = new Map();
          combinedData.set("Customer Name", salesOrder.customer_name);
          combinedData.set("SO Date", salesOrder.date);
          combinedData.set("SO Number", salesOrder.salesorder_number);
          combinedData.set("Customer PO", salesOrder.reference_number);
          combinedData.set("Amount", salesOrder.total);
          combinedData.set("Supplier PO No", " ");
          combinedData.set("Supplier Name", " ");
          combinedData.set("Invoice Number", invoice.invoice_number);
          combinedData.set("Invoice Value", " ");
          combinedData.set("Bill Number", " ");
          combinedData.set("Bill Value", " ");
          resultArray.push(combinedData);
        });

        if (relevantInvoices.length === 0) {
          const combinedData = new Map();
          combinedData.set("Customer Name", salesOrder.customer_name);
          combinedData.set("SO Date", salesOrder.date);
          combinedData.set("SO Number", salesOrder.salesorder_number);
          combinedData.set("Customer PO", salesOrder.reference_number);
          combinedData.set("Amount", salesOrder.total);
          combinedData.set("Supplier PO No", " ");
          combinedData.set("Supplier Name", " ");
          combinedData.set("Invoice Number", " ");
          combinedData.set("Invoice Value", " ");
          combinedData.set("Bill Number", " ");
          combinedData.set("Bill Value", " ");
          resultArray.push(combinedData);
        }
      }
      // const relevantbills = billsAll.filter(bill => bill.reference_number === salesOrder.salesorder_number);
      // relevantPurchaseOrders.forEach(purchaseOrder => {
      //   relevantinvoices.forEach(invoice => {
      //     const combinedData = new Map();
      //     combinedData.set("Customer Name", salesOrder.customer_name);
      //     combinedData.set("SO Date", salesOrder.date);
      //     combinedData.set("SO Number", salesOrder.salesorder_number);
      //     combinedData.set("Customer PO", salesOrder.reference_number);
      //     combinedData.set("Amount", salesOrder.total);
      //     combinedData.set("Supplier PO No", purchaseOrder.purchaseorder_number);
      //     combinedData.set("Supplier Name", purchaseOrder.vendor_name);
      //     combinedData.set("Supplier PO Value", purchaseOrder.total);
      //     combinedData.set("Invoice Number", invoice.invoice_number);
      //     combinedData.set("Invoice Value", invoice.total);
      //     resultArray.push(combinedData);
      //   });


      //   if (relevantinvoices.length === 0) {
      //     const combinedData = new Map();
      //     combinedData.set("Customer Name", salesOrder.customer_name);
      //     combinedData.set("SO Date", salesOrder.date);
      //     combinedData.set("SO Number", salesOrder.salesorder_number);
      //     combinedData.set("Customer PO", salesOrder.reference_number);
      //     combinedData.set("Amount", salesOrder.total);
      //     combinedData.set("Supplier PO No", purchaseOrder.purchaseorder_number);
      //     combinedData.set("Supplier Name", purchaseOrder.vendor_name);
      //     combinedData.set("Supplier PO Value", purchaseOrder.total);
      //     combinedData.set("Invoice Number", " ");
      //     combinedData.set("Invoice Value", " ");
      //     resultArray.push(combinedData);
      //   }

      // });

      

      // if (relevantPurchaseOrders.length === 0) {
      //   relevantinvoices.forEach(invoice => {
      //     const combinedData = new Map();
      //     combinedData.set("Customer Name", salesOrder.customer_name);
      //     combinedData.set("SO Date", salesOrder.date);
      //     combinedData.set("SO Number", salesOrder.salesorder_number);
      //     combinedData.set("Customer PO", salesOrder.reference_number);
      //     combinedData.set("Amount", salesOrder.total);
      //     combinedData.set("Supplier PO No", " ");
      //     combinedData.set("Supplier Name", " ");
      //     combinedData.set("Supplier PO Value", " ");
      //     combinedData.set("Invoice Number", invoice.invoice_number);
      //     combinedData.set("Invoice Value", invoice.total);
      //     resultArray.push(combinedData);
      //   });

      //   if (relevantinvoices.length === 0) {
      //     const combinedData = new Map();
      //     combinedData.set("Customer Name", salesOrder.customer_name);
      //     combinedData.set("SO Date", salesOrder.date);
      //     combinedData.set("SO Number", salesOrder.salesorder_number);
      //     combinedData.set("Customer PO", salesOrder.reference_number);
      //     combinedData.set("Amount", salesOrder.total);
      //     combinedData.set("Supplier PO No", " ");
      //     combinedData.set("Supplier Name", " ");
      //     combinedData.set("Supplier PO Value", " ");
      //     combinedData.set("Invoice Number", " ");
      //     combinedData.set("Invoice Value", " ");
      //     resultArray.push(combinedData);
      //   }
      // }
    });

    await insertDataIntoSheet(accessToken, resultArray);
  } catch (error) {
    console.error('Error:', error);
  }
}

main();
