export default async function handler(req, res) {
  if (req.method !== "GET") {
    return res.status(405).json({ status: "error", message: "Method not allowed" });
  }

  const uid = (req.query.uid || "").toString().trim();
  if (!uid) {
    return res.status(400).json({ status: "error", message: "UID kosong" });
  }

  const scriptUrl = process.env.GOOGLE_SCRIPT_URL;
  if (!scriptUrl) {
    return res.status(500).json({
      status: "error",
      message: "Environment variable GOOGLE_SCRIPT_URL belum diatur di Vercel"
    });
  }

  try {
    const response = await fetch(`${scriptUrl}?uid=${encodeURIComponent(uid)}`, {
      method: "GET",
      redirect: "follow"
    });

    const text = await response.text();

    try {
      const data = JSON.parse(text);
      return res.status(200).json(data);
    } catch {
      return res.status(502).json({
        status: "error",
        message: "Respons Apps Script bukan JSON valid (cek akses Web App: Anyone)"
      });
    }
  } catch (err) {
    return res.status(500).json({
      status: "error",
      message: "Gagal menghubungi Google Apps Script"
    });
  }
}
