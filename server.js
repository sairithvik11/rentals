require('dotenv').config();
const express = require('express');
const cors = require('cors');
const fs = require('fs');
const path = require('path');
const nodemailer = require('nodemailer');

const XLSX = require('xlsx');

const app = express();
const PORT = process.env.PORT || 3000;
const DATA_FILE = path.join(__dirname, 'vehicles.json');
const BOOKINGS_FILE = path.join(__dirname, 'bookings.json');
const EXCEL_FILE = path.join(__dirname, 'bookings.xlsx');
const ENV_FILE = path.join(__dirname, '.env');

app.use(cors());
app.use(express.json({ limit: '10mb' }));
app.use(express.static(__dirname));

app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});


function getVehicles() {
  try {
    if (fs.existsSync(DATA_FILE)) {
      const raw = fs.readFileSync(DATA_FILE, 'utf8');
      return JSON.parse(raw);
    }
  } catch (err) {
    console.error('Error reading vehicles.json:', err);
  }
  return [];
}

function saveVehicles(vehicles) {
  try {
    fs.writeFileSync(DATA_FILE, JSON.stringify(vehicles, null, 2), 'utf8');
  } catch (err) {
    console.error('Error saving vehicles.json:', err);
  }
}

function getBookings() {
  try {
    if (fs.existsSync(BOOKINGS_FILE)) {
      const raw = fs.readFileSync(BOOKINGS_FILE, 'utf8');
      return JSON.parse(raw);
    }
  } catch (err) {
    console.error('Error reading bookings.json:', err);
  }
  return [];
}

function saveBookingsToExcel(bookings) {
  try {
    const excelData = (bookings || []).map(b => ({
      'Booking Code': b.bookingCode || '',
      'Date & Time': b.createdAt ? new Date(b.createdAt).toLocaleString('en-IN') : '',
      'Customer Name': (b.user && b.user.name) || 'Explorer',
      'Customer Email': (b.user && b.user.email) || '',
      'Customer Phone': (b.user && b.user.phone) || '',
      'Vehicle': b.vehicle ? `${b.vehicle.brand} ${b.vehicle.model}` : '',
      'Pickup Date': b.startDate || '',
      'Return Date': b.endDate || '',
      'Pickup Location': b.location || '',
      'Days': b.days || 1,
      'Total Amount (₹)': b.totalAmount || 0,
      'Payment Method': b.paymentMethod || 'Pay at Pickup',
      'Payment Status': b.paymentStatus || 'Paid',
      'Status': b.status || 'Confirmed',
      'Email Sent': b.emailSent ? 'Yes' : 'No'
    }));

    const worksheet = XLSX.utils.json_to_sheet(excelData);
    worksheet['!cols'] = [
      { wch: 18 }, { wch: 22 }, { wch: 18 }, { wch: 26 }, { wch: 15 },
      { wch: 24 }, { wch: 14 }, { wch: 14 }, { wch: 24 }, { wch: 8 },
      { wch: 18 }, { wch: 20 }, { wch: 16 }, { wch: 12 }, { wch: 14 }
    ];

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Bookings');
    XLSX.writeFile(workbook, EXCEL_FILE);
    console.log(`📊 Updated bookings.xlsx spreadsheet`);
  } catch (err) {
    console.error('Error saving bookings.xlsx:', err);
  }
}

function saveBookings(bookings) {
  try {
    fs.writeFileSync(BOOKINGS_FILE, JSON.stringify(bookings, null, 2), 'utf8');
    saveBookingsToExcel(bookings);
  } catch (err) {
    console.error('Error saving bookings.json:', err);
  }
}

let transporter = null;
let isRealSMTP = false;

function createEmailTransporter() {
  const smtpUser = (process.env.SMTP_USER || '').trim();
  const smtpPass = (process.env.SMTP_PASS || '').trim().replace(/\s+/g, '');
  const smtpHost = (process.env.SMTP_HOST || 'smtp.gmail.com').trim();
  const smtpPort = parseInt(process.env.SMTP_PORT || '465');

  if (smtpUser && smtpPass && smtpUser !== '' && smtpPass !== '') {
    transporter = nodemailer.createTransport({
      host: smtpHost,
      port: smtpPort,
      secure: smtpPort === 465,
      auth: { user: smtpUser, pass: smtpPass }
    });
    isRealSMTP = true;
    console.log(`✅ Real SMTP Transporter active using: ${smtpUser} (${smtpHost}:${smtpPort})`);
  } else {
    isRealSMTP = false;
  }
}

async function initEmailTransporter() {
  createEmailTransporter();

  if (!isRealSMTP) {
    try {
      const testAccount = await nodemailer.createTestAccount();
      transporter = nodemailer.createTransport({
        host: 'smtp.ethereal.email',
        port: 587,
        secure: false,
        auth: {
          user: testAccount.user,
          pass: testAccount.pass
        }
      });
      console.log(`ℹ️ Real SMTP credentials not in .env — using Ethereal sandbox test account (${testAccount.user}).`);
    } catch (err) {
      console.error('Fallback transport initialized');
      transporter = nodemailer.createTransport({ jsonTransport: true });
    }
  }
}
initEmailTransporter();

/* ============================================================
   API ENDPOINTS
============================================================ */

// GET /api/vehicles - Fetch all vehicles
app.get('/api/vehicles', (req, res) => {
  const vehicles = getVehicles();
  res.json({ success: true, vehicles });
});

// POST /api/vehicles - Add a new vehicle
app.post('/api/vehicles', (req, res) => {
  const newVehicle = req.body;
  if (!newVehicle || !newVehicle.brand || !newVehicle.model) {
    return res.status(400).json({ success: false, message: 'Invalid vehicle details' });
  }

  const vehicles = getVehicles();
  newVehicle.id = newVehicle.id || 'v_' + Date.now();
  vehicles.unshift(newVehicle);
  saveVehicles(vehicles);

  console.log(`🚗 Vehicle added: ${newVehicle.brand} ${newVehicle.model} (${newVehicle.id})`);
  res.json({ success: true, message: 'Vehicle added successfully', vehicle: newVehicle, vehicles });
});

// DELETE /api/vehicles/:id - Delete a vehicle
app.delete('/api/vehicles/:id', (req, res) => {
  const { id } = req.params;
  let vehicles = getVehicles();
  const initialLength = vehicles.length;
  vehicles = vehicles.filter(v => v.id !== id);

  if (vehicles.length === initialLength) {
    return res.status(404).json({ success: false, message: 'Vehicle not found' });
  }

  saveVehicles(vehicles);
  console.log(`🗑️ Vehicle deleted: ${id}`);
  res.json({ success: true, message: 'Vehicle deleted successfully', id, vehicles });
});

// GET /api/config/smtp - Check SMTP status
app.get('/api/config/smtp', (req, res) => {
  res.json({
    success: true,
    isRealSMTP,
    smtpUser: process.env.SMTP_USER || '',
    smtpHost: process.env.SMTP_HOST || 'smtp.gmail.com',
    smtpPort: process.env.SMTP_PORT || '465'
  });
});

// POST /api/config/smtp - Update SMTP configuration dynamically
app.post('/api/config/smtp', (req, res) => {
  const { user, pass, host, port } = req.body;
  if (!user || !pass) {
    return res.status(400).json({ success: false, message: 'Email address and App Password are required' });
  }

  process.env.SMTP_USER = user.trim();
  process.env.SMTP_PASS = pass.trim();
  process.env.SMTP_HOST = (host || 'smtp.gmail.com').trim();
  process.env.SMTP_PORT = (port || '465').toString().trim();

  const envContent = `SMTP_USER=${process.env.SMTP_USER}\nSMTP_PASS=${process.env.SMTP_PASS}\nSMTP_HOST=${process.env.SMTP_HOST}\nSMTP_PORT=${process.env.SMTP_PORT}\n`;
  fs.writeFileSync(ENV_FILE, envContent, 'utf8');

  createEmailTransporter();

  console.log(`🔑 Real SMTP configuration updated for ${user.trim()}`);
  res.json({
    success: true,
    message: `SMTP successfully configured for ${user.trim()}! Real emails will now land directly in your inbox.`,
    isRealSMTP: true,
    smtpUser: process.env.SMTP_USER
  });
});

// GET /api/bookings - Fetch all bookings
app.get('/api/bookings', (req, res) => {
  const bookings = getBookings();
  res.json({ success: true, bookings });
});

// GET /api/bookings/export - Download Excel spreadsheet of bookings
app.get('/api/bookings/export', (req, res) => {
  if (fs.existsSync(EXCEL_FILE)) {
    res.download(EXCEL_FILE, 'bookings.xlsx');
  } else {
    saveBookingsToExcel(getBookings());
    if (fs.existsSync(EXCEL_FILE)) {
      res.download(EXCEL_FILE, 'bookings.xlsx');
    } else {
      res.status(404).json({ success: false, message: 'Spreadsheet file not found' });
    }
  }
});

// POST /api/bookings - Process booking & send email
app.post('/api/bookings', async (req, res) => {
  const { vehicle, user, startDate, endDate, location, totalAmount, days, paymentMethod, paymentStatus } = req.body;

  if (!vehicle || !user || !user.email) {
    return res.status(400).json({ success: false, message: 'Vehicle and User Email are required' });
  }

  const methodStr = paymentMethod || 'Pay at Pickup';
  const statusStr = paymentStatus || (methodStr === 'Pay at Pickup' ? 'Pending at Pickup' : 'Paid');

  const bookingCode = 'RIDEMITRA-' + Math.random().toString(36).slice(2, 8).toUpperCase();

  const emailHtml = `
    <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; max-width: 600px; margin: 0 auto; background: #0F172A; color: #F8FAFC; border-radius: 16px; overflow: hidden; border: 1px solid #334155;">
      <div style="background: linear-gradient(135deg, #FF9500 0%, #E68600 100%); padding: 28px 32px; text-align: center;">
        <h1 style="margin: 0; color: #FFFFFF; font-size: 26px; font-weight: 800; letter-spacing: 0.5px;">RideMitra</h1>
        <p style="margin: 4px 0 0 0; color: rgba(255,255,255,0.9); font-size: 14px;">Offroad & Luxury Car Rentals</p>
      </div>

      <div style="padding: 32px;">
        <h2 style="color: #FF9500; font-size: 20px; margin-top: 0;">🎉 Booking Confirmed!</h2>
        <p style="color: #94A3B8; font-size: 15px; line-height: 1.6;">Hi <strong>${user.name || 'Explorer'}</strong>,</p>
        <p style="color: #CBD5E1; font-size: 15px; line-height: 1.6;">Your booking for <strong>${vehicle.brand} ${vehicle.model} (${vehicle.year})</strong> is locked in and ready for pickup!</p>

        <div style="background: #1E293B; border-radius: 12px; padding: 20px; margin: 24px 0; border: 1px solid #334155;">
          <div style="font-size: 12px; text-transform: uppercase; color: #94A3B8; letter-spacing: 1px; margin-bottom: 8px;">Booking Reference</div>
          <div style="font-size: 22px; font-weight: 800; color: #FF9500; letter-spacing: 2px;">${bookingCode}</div>
        </div>

        <table style="width: 100%; border-collapse: collapse; margin-bottom: 24px;">
          <tr>
            <td style="padding: 10px 0; color: #94A3B8; border-bottom: 1px solid #334155;">Vehicle:</td>
            <td style="padding: 10px 0; color: #F8FAFC; font-weight: 600; text-align: right; border-bottom: 1px solid #334155;">${vehicle.brand} ${vehicle.model}</td>
          </tr>
          <tr>
            <td style="padding: 10px 0; color: #94A3B8; border-bottom: 1px solid #334155;">Pickup Date:</td>
            <td style="padding: 10px 0; color: #F8FAFC; font-weight: 600; text-align: right; border-bottom: 1px solid #334155;">${startDate}</td>
          </tr>
          <tr>
            <td style="padding: 10px 0; color: #94A3B8; border-bottom: 1px solid #334155;">Return Date:</td>
            <td style="padding: 10px 0; color: #F8FAFC; font-weight: 600; text-align: right; border-bottom: 1px solid #334155;">${endDate} (${days || 1} days)</td>
          </tr>
          <tr>
            <td style="padding: 10px 0; color: #94A3B8; border-bottom: 1px solid #334155;">Pickup Location:</td>
            <td style="padding: 10px 0; color: #F8FAFC; font-weight: 600; text-align: right; border-bottom: 1px solid #334155;">${location || 'Central Hub'}</td>
          </tr>
          <tr>
            <td style="padding: 10px 0; color: #94A3B8; border-bottom: 1px solid #334155;">Payment Method:</td>
            <td style="padding: 10px 0; color: #F8FAFC; font-weight: 600; text-align: right; border-bottom: 1px solid #334155;">${methodStr}</td>
          </tr>
          <tr>
            <td style="padding: 10px 0; color: #94A3B8; border-bottom: 1px solid #334155;">Payment Status:</td>
            <td style="padding: 10px 0; color: #10B981; font-weight: 700; text-align: right; border-bottom: 1px solid #334155;">${statusStr}</td>
          </tr>
          <tr>
            <td style="padding: 12px 0; color: #F8FAFC; font-weight: 700; font-size: 16px;">Total Amount:</td>
            <td style="padding: 12px 0; color: #FF9500; font-weight: 800; font-size: 18px; text-align: right;">₹${(totalAmount || 0).toLocaleString('en-IN')}</td>
          </tr>
        </table>

        <p style="color: #94A3B8; font-size: 13.5px; line-height: 1.5;">Please present this booking code <strong>${bookingCode}</strong> along with your valid driving license at the pickup location.</p>
      </div>

      <div style="background: #020617; padding: 20px; text-align: center; color: #64748B; font-size: 12px;">
        © 2026 RideMitra Rentals. All rights reserved.<br>Need help? Contact support@ridemitra.com
      </div>
    </div>
  `;

  let emailPreviewUrl = null;
  let emailSent = false;
  let smtpError = null;

  if (transporter) {
    try {
      const fromAddr = (isRealSMTP && process.env.SMTP_USER) ? `"RideMitra" <${process.env.SMTP_USER}>` : '"RideMitra Bookings" <no-reply@ridemitra.com>';
      const info = await transporter.sendMail({
        from: fromAddr,
        to: user.email,
        subject: `✨ Booking Confirmed [${bookingCode}]: ${vehicle.brand} ${vehicle.model}`,
        html: emailHtml
      });

      emailSent = true;
      if (!isRealSMTP) {
        emailPreviewUrl = nodemailer.getTestMessageUrl(info);
      }
      console.log(`✉️ Booking email sent to ${user.email}! (Real SMTP delivery: ${isRealSMTP})`);
    } catch (err) {
      console.error('Error sending email via Nodemailer:', err);
      smtpError = err.message || 'SMTP Authentication failed';
    }
  }

  const newBookingRecord = {
    bookingCode,
    createdAt: new Date().toISOString(),
    vehicle: {
      id: vehicle.id,
      brand: vehicle.brand,
      model: vehicle.model,
      year: vehicle.year,
      price: vehicle.price,
      type: vehicle.type
    },
    user: {
      name: user.name || 'Explorer',
      email: user.email,
      phone: user.phone || ''
    },
    startDate: startDate || '',
    endDate: endDate || '',
    location: location || 'Central Hub',
    totalAmount: totalAmount || 0,
    days: days || 1,
    paymentMethod: methodStr,
    paymentStatus: statusStr,
    status: 'Confirmed',
    emailSent,
    isRealSMTP
  };

  const bookings = getBookings();
  bookings.unshift(newBookingRecord);
  saveBookings(bookings);
  console.log(`💾 Booking saved to bookings.json: ${bookingCode}`);

  res.json({
    success: true,
    message: isRealSMTP && emailSent ? `Booking confirmed and real email delivered to ${user.email}!` : 'Booking confirmed!',
    bookingCode,
    emailSent,
    isRealSMTP,
    smtpError,
    emailPreviewUrl,
    emailHtml,
    bookingDetails: {
      vehicle: `${vehicle.brand} ${vehicle.model}`,
      user: user.name,
      email: user.email,
      totalAmount
    }
  });
});

// POST /api/bookings/cancel - Cancel a booking & send confirmation email
app.post('/api/bookings/cancel', async (req, res) => {
  const { bookingCode, reason, userEmail } = req.body;
  if (!bookingCode) {
    return res.status(400).json({ success: false, message: 'Booking Code is required' });
  }

  const bookings = getBookings();
  const booking = bookings.find(b => b.bookingCode === bookingCode);

  if (!booking) {
    return res.status(404).json({ success: false, message: 'Booking not found' });
  }

  booking.status = 'Cancelled';
  booking.cancelReason = reason || 'Cancelled by customer';
  booking.cancelledAt = new Date().toISOString();

  saveBookings(bookings);
  console.log(`❌ Booking cancelled: ${bookingCode}`);

  const emailToUse = userEmail || (booking.user && booking.user.email);
  let emailSent = false;
  let smtpError = null;

  if (transporter && emailToUse) {
    const cancelEmailHtml = `
      <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; max-width: 600px; margin: 0 auto; background: #0F172A; color: #F8FAFC; border-radius: 16px; overflow: hidden; border: 1px solid #334155;">
        <div style="background: linear-gradient(135deg, #EF4444 0%, #DC2626 100%); padding: 28px 32px; text-align: center;">
          <h1 style="margin: 0; color: #FFFFFF; font-size: 26px; font-weight: 800;">RideMitra</h1>
          <p style="margin: 4px 0 0 0; color: rgba(255,255,255,0.9); font-size: 14px;">Booking Cancellation Notice</p>
        </div>

        <div style="padding: 32px;">
          <h2 style="color: #EF4444; font-size: 20px; margin-top: 0;">❌ Booking Cancelled</h2>
          <p style="color: #94A3B8; font-size: 15px; line-height: 1.6;">Hi <strong>${(booking.user && booking.user.name) || 'Explorer'}</strong>,</p>
          <p style="color: #CBD5E1; font-size: 15px; line-height: 1.6;">Your booking for <strong>${booking.vehicle ? booking.vehicle.brand + ' ' + booking.vehicle.model : 'Vehicle'}</strong> has been successfully cancelled as requested.</p>

          <div style="background: #1E293B; border-radius: 12px; padding: 20px; margin: 24px 0; border: 1px solid #334155;">
            <div style="font-size: 12px; text-transform: uppercase; color: #94A3B8; letter-spacing: 1px; margin-bottom: 8px;">Cancelled Reference</div>
            <div style="font-size: 22px; font-weight: 800; color: #EF4444; letter-spacing: 2px;">${bookingCode}</div>
          </div>

          <table style="width: 100%; border-collapse: collapse; margin-bottom: 24px;">
            <tr>
              <td style="padding: 10px 0; color: #94A3B8; border-bottom: 1px solid #334155;">Vehicle:</td>
              <td style="padding: 10px 0; color: #F8FAFC; font-weight: 600; text-align: right; border-bottom: 1px solid #334155;">${booking.vehicle ? booking.vehicle.brand + ' ' + booking.vehicle.model : ''}</td>
            </tr>
            <tr>
              <td style="padding: 10px 0; color: #94A3B8; border-bottom: 1px solid #334155;">Dates:</td>
              <td style="padding: 10px 0; color: #F8FAFC; font-weight: 600; text-align: right; border-bottom: 1px solid #334155;">${booking.startDate || ''} to ${booking.endDate || ''}</td>
            </tr>
            <tr>
              <td style="padding: 10px 0; color: #94A3B8; border-bottom: 1px solid #334155;">Status:</td>
              <td style="padding: 10px 0; color: #EF4444; font-weight: 800; text-align: right; border-bottom: 1px solid #334155;">CANCELLED</td>
            </tr>
            <tr>
              <td style="padding: 12px 0; color: #F8FAFC; font-weight: 700; font-size: 16px;">Refund Status:</td>
              <td style="padding: 12px 0; color: #10B981; font-weight: 800; font-size: 16px; text-align: right;">Full Refund Initiated (₹${(booking.totalAmount || 0).toLocaleString('en-IN')})</td>
            </tr>
          </table>

          <p style="color: #94A3B8; font-size: 13.5px; line-height: 1.5;">If you have any questions regarding your cancellation or refund, please reach out to support@ridemitra.com.</p>
        </div>

        <div style="background: #020617; padding: 20px; text-align: center; color: #64748B; font-size: 12px;">
          © 2026 RideMitra Rentals. All rights reserved.
        </div>
      </div>
    `;

    try {
      const fromAddr = (isRealSMTP && process.env.SMTP_USER) ? `"RideMitra" <${process.env.SMTP_USER}>` : '"RideMitra Bookings" <no-reply@ridemitra.com>';
      await transporter.sendMail({
        from: fromAddr,
        to: emailToUse,
        subject: `❌ Booking Cancelled [${bookingCode}]: ${booking.vehicle ? booking.vehicle.brand + ' ' + booking.vehicle.model : 'Vehicle'}`,
        html: cancelEmailHtml
      });
      emailSent = true;
      console.log(`✉️ Cancellation email sent to ${emailToUse}`);
    } catch (err) {
      console.error('Error sending cancellation email:', err);
      smtpError = err.message;
    }
  }

  res.json({
    success: true,
    message: `Booking ${bookingCode} cancelled successfully.${emailSent ? ' Cancellation email sent.' : ''}`,
    booking,
    emailSent,
    bookings
  });
});

// POST /api/vehicles/rate - Rate a vehicle and submit review
app.post('/api/vehicles/rate', (req, res) => {
  const { vehicleId, rating, review, userName } = req.body;
  if (!vehicleId || !rating) {
    return res.status(400).json({ success: false, message: 'Vehicle ID and rating (1-5) are required' });
  }

  const vehicles = getVehicles();
  let vehicle = vehicles.find(v => v.id === vehicleId);

  if (!vehicle) {
    vehicle = { id: vehicleId, brand: 'Vehicle', model: '' };
    vehicles.push(vehicle);
  }

  vehicle.reviewsList = vehicle.reviewsList || [];
  vehicle.reviewsList.push({
    rating: parseInt(rating),
    review: review || '',
    userName: userName || 'Rider',
    createdAt: new Date().toISOString()
  });

  const total = vehicle.reviewsList.reduce((acc, curr) => acc + curr.rating, 0);
  vehicle.avgRating = parseFloat((total / vehicle.reviewsList.length).toFixed(1));
  vehicle.ratingCount = vehicle.reviewsList.length;

  saveVehicles(vehicles);
  console.log(`⭐ New review for vehicle ${vehicleId}: ${rating} stars by ${userName || 'Rider'}`);

  res.json({
    success: true,
    message: 'Thank you for rating your experience!',
    vehicle,
    vehicles
  });
});

// Auto-sync Excel file on boot
saveBookingsToExcel(getBookings());

if (require.main === module) {
  const server = app.listen(PORT, () => {
    console.log(`🚀 RideMitra Server running on http://localhost:${PORT}`);
  });

  server.on('error', (err) => {
    if (err.code === 'EADDRINUSE') {
      console.error(`\n⚠️ Port ${PORT} is already in use by another running server instance.`);
      console.error(`👉 RideMitra is ALREADY running on http://localhost:${PORT}!\n`);
    } else {
      console.error('Server error:', err);
    }
  });
}

module.exports = app;

