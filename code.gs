function myFunction() {
	Logger.log(MailApp.getRemainingDailyQuota());

	for (let i = 2; i <= 40; i++) {
		sendEmailByRow(i);
	}
}

function sendEmailByRow(row) {
	const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FormResponses");
	if (!sheet) throw new Error('Sheet "FormResponses" not found');

	row = Number(row);
	if (!row || row < 2) throw new Error('Invalid row number');

	const statusCol = 17;

	try {
		if (MailApp.getRemainingDailyQuota() < 1) throw new Error('No email quota remaining today');

		const emailAddress = String(sheet.getRange(row, 2).getValue() || "").trim();
		const name         = String(sheet.getRange(row, 3).getValue() || "").trim();

		if (!emailAddress) {
			sheet.getRange(row, statusCol).setValue('Skipped: no email');
			return;
		}
		if (!name) throw new Error('Missing name in column C');

		const emailData = { fullName: name };

		const subject = "ประกาศผลการคัดเลือก Staff — ประชุมผู้ปกครอง ม.4 TU89";
		const plainText = generatePlainText(emailData);
		const htmlBody  = generateHtmlBody(emailData);

		GmailApp.sendEmail(emailAddress, subject, plainText, {
			htmlBody: htmlBody,
			name: "Triamudom Family",
			noReply: false
		});

		sheet.getRange(row, statusCol).setValue('Email Sent Successfully');
		Logger.log("Email sent successfully for row " + row);
	} catch (err) {
		sheet.getRange(row, statusCol).setValue("Sending Failed: " + err.message);
		Logger.log("Sending failed for row " + row + ": " + err.message);
		throw err;
	}
}

function escapeHtml_(text) {
	return String(text ?? '')
		.replace(/&/g, '&amp;')
		.replace(/</g, '&lt;')
		.replace(/>/g, '&gt;')
		.replace(/"/g, '&quot;')
		.replace(/'/g, '&#39;');
}

function generatePlainText(data) {
	const fullName = data.fullName || data.name || "";

	return `
ประกาศผลการคัดเลือก

ผ่านการคัดเลือก

${fullName}

ผ่านการคัดเลือกเข้าเป็น Staff ในการประชุมผู้ปกครอง ม.4 TU89

การยืนยันสิทธิ์เข้าร่วมเป็น Staff

กรุณาเข้าร่วมกลุ่มไลน์ภายในวันนี้ เพื่อยืนยันสิทธิ์
และรับข้อมูลเพิ่มเติมเกี่ยวกับการชี้แจงรายละเอียดงาน

เข้ากลุ่มไลน์เพื่อยืนยันสิทธิ์:
https://line.me/ti/g/7AC7heEEgY

กำหนดการสำคัญ

Briefing (Zoom)
วันที่: 5 พฤษภาคม 2569
เวลา: 18:00 – 19:00 น.

วันทำงาน
วันที่: 16 พฤษภาคม 2569
เวลา: 07:00 – 15:00 น. (Full day)

Zoom Meeting

Meeting ID: 822 2580 7426
Passcode: 964836

ลิงก์เข้าร่วม Zoom Meeting:
https://us06web.zoom.us/j/82225807426?pwd=eDFq5uw6jciUjVIs8zjaEJzpp7p40n.1

ติดต่อสอบถามเพิ่มเติม

Line Official Account: @triamudom
โทรติดต่อ: 0803323330

ขอแสดงความนับถือ

คณะทำงาน Triamudom Family
สมาคมผู้ปกครองและครูโรงเรียนเตรียมอุดมศึกษา
	`.trim();
}

function generateHtmlBody(data) {
	return `
<!DOCTYPE html>
<html lang="th">

<head>
	<meta charset="UTF-8" />
	<meta name="viewport" content="width=device-width, initial-scale=1.0" />
	<link rel="preconnect" href="https://fonts.googleapis.com" />
	<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin />
	<link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Sans+Thai:wght@400;500;600;700&display=swap"
		rel="stylesheet" />
	<title>Staff Email Announcement</title>
</head>

<body
	style="margin:0;box-sizing:border-box;width:100%;-webkit-text-size-adjust:100%;font-family:'IBM Plex Sans Thai','Noto Sans Thai','Anuphan','Kanit','Tahoma',Arial,sans-serif;font-weight:400;background:linear-gradient(135deg,#fff1f2,#fce7f3,#fae8ff);min-height:100vh;padding:40px 24px;color:#1f2937;">
	<div
		style="max-width:768px;margin:0 auto;overflow:hidden;border-radius:24px;border:1px solid rgba(255,255,255,0.9);background:rgba(255,255,255,0.92);box-shadow:0 20px 35px rgba(236,72,153,0.14),0 8px 14px rgba(236,72,153,0.08);width:100%;box-sizing:border-box;font-family:'IBM Plex Sans Thai','Noto Sans Thai','Anuphan','Kanit','Tahoma',Arial,sans-serif;font-weight:400;">

		<div
			style="background:linear-gradient(to top right,#f472b6,#fb7185,#f43f5e);padding:40px 48px;color:#ffffff;text-align:center;">
			<div
				style="margin-bottom:16px;display:inline-block;border-radius:9999px;background-color:rgba(255,255,255,0.15);border:1px solid rgba(255,255,255,0.30);padding:4px 16px;font-size:12px;line-height:16px;font-weight:400;">
				ประกาศผลการคัดเลือก
			</div>

			<h1 style="margin:0;font-size:36px;line-height:40px;font-weight:700;letter-spacing:-0.025em;">
				ผ่านการคัดเลือก
			</h1>

			<h2 style="margin:8px 0 0 0;font-size:30px;line-height:36px;font-weight:600;opacity:0.9;">
				${escapeHtml_(data.fullName)}
			</h2>

			<div
				style="margin-top:20px;display:inline-block;border-radius:9999px;background-color:rgba(255,255,255,0.15);border:1px solid rgba(255,255,255,0.30);padding:8px 16px;font-size:14px;line-height:20px;font-weight:400;">
				ผ่านการคัดเลือกเข้าเป็น Staff ในการประชุมผู้ปกครอง ม.4 TU89
			</div>
		</div>

		<div style="padding:32px 48px;">

			<section style="margin:0;">
				<h3
					style="margin:0 0 16px 0;font-size:12px;line-height:16px;font-weight:600;letter-spacing:1.2px;color:#f472b6;text-transform:uppercase;">
					การยืนยันสิทธิ์เข้าร่วมเป็น Staff
				</h3>

				<p style="margin:0;font-size:16px;line-height:28px;color:#374151;font-weight:400;">
					กรุณาเข้าร่วมกลุ่มไลน์ภายวันนี้ เพื่อยืนยันสิทธิ์
					และรับข้อมูลเพิ่มเติมเกี่ยวกับการชี้แจงรายละเอียดงาน
				</p>

				<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"
					style="margin-top:16px;">
					<tr>
						<td width="50%" valign="top" style="padding-right:6px;">
							<table role="presentation" width="100%" height="100" cellpadding="0" cellspacing="0"
								border="0"
								style="height:100px;border:1px solid #fce7f3;background-color:#fdf2f8;border-radius:12px;border-collapse:separate;">
								<tr>
									<td valign="top" style="padding:14px 16px;">
										<p
											style="margin:0 0 4px 0;font-size:12px;line-height:16px;color:#f472b6;font-weight:500;">
											Briefing (Zoom)
										</p>
										<p
											style="margin:0;font-size:14px;line-height:20px;font-weight:600;color:#9f1239;">
											5 พฤษภาคม 2569
										</p>
										<p
											style="margin:2px 0 0 0;font-size:12px;line-height:16px;color:#4b5563;font-weight:400;">
											18:00 – 19:00 น.
										</p>
									</td>
								</tr>
							</table>
						</td>

						<td width="50%" valign="top" style="padding-left:6px;">
							<table role="presentation" width="100%" height="100" cellpadding="0" cellspacing="0"
								border="0"
								style="height:100px;border:1px solid #fce7f3;background-color:#fdf2f8;border-radius:12px;border-collapse:separate;">
								<tr>
									<td valign="top" style="padding:14px 16px;">
										<p
											style="margin:0 0 4px 0;font-size:12px;line-height:16px;color:#f472b6;font-weight:500;">
											วันทำงาน
										</p>
										<p
											style="margin:0;font-size:14px;line-height:20px;font-weight:600;color:#9f1239;">
											16 พฤษภาคม 2569
										</p>
										<p
											style="margin:2px 0 0 0;font-size:12px;line-height:16px;color:#4b5563;font-weight:400;">
											07:00 – 15:00 น. (Full day)
										</p>
									</td>
								</tr>
							</table>
						</td>
					</tr>
				</table>

				<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"
					style="margin-top:16px;">
					<tr>
						<td align="center" style="text-align:center;">
							<a href="https://line.me/ti/g/7AC7heEEgY" target="_blank" rel="noopener noreferrer"
								style="display:inline-block;text-decoration:none;color:#ffffff;background:linear-gradient(135deg,#22c55e,#16a34a);border-radius:12px;padding:12px 20px;font-size:14px;line-height:20px;font-weight:600;font-family:'IBM Plex Sans Thai','Noto Sans Thai','Anuphan','Kanit','Tahoma',Arial,sans-serif;">
								เข้ากลุ่มไลน์เพื่อยืนยันสิทธิ์
							</a>
						</td>
					</tr>
				</table>
			</section>

			<section style="margin-top:32px;">
				<h3
					style="margin:0 0 16px 0;font-size:12px;line-height:16px;font-weight:600;letter-spacing:1.2px;color:#60a5fa;text-transform:uppercase;">
					Zoom Meeting
				</h3>

				<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"
					style="border:1px solid #dbeafe;background-color:#eff6ff;border-radius:16px;border-collapse:separate;overflow:hidden;">
					<tr>
						<td style="padding:16px 20px;border-bottom:1px solid #dbeafe;">
							<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0">
								<tr>
									<td width="50%" valign="top"
										style="padding-right:16px;border-right:1px solid #dbeafe;">
										<p
											style="margin:0 0 2px 0;font-size:12px;line-height:16px;color:#60a5fa;font-weight:500;">
											Meeting ID
										</p>
										<p
											style="margin:0;font-size:14px;line-height:20px;font-weight:600;color:#1f2937;letter-spacing:0.025em;">
											822 2580 7426
										</p>
									</td>

									<td width="50%" valign="top" style="padding-left:16px;">
										<p
											style="margin:0 0 2px 0;font-size:12px;line-height:16px;color:#60a5fa;font-weight:500;">
											Passcode
										</p>
										<p
											style="margin:0;font-size:14px;line-height:20px;font-weight:600;color:#1f2937;letter-spacing:0.025em;">
											964836
										</p>
									</td>
								</tr>
							</table>
						</td>
					</tr>

					<tr>
						<td align="center" style="padding:16px 20px;text-align:center;">
							<a href="https://us06web.zoom.us/j/82225807426?pwd=eDFq5uw6jciUjVIs8zjaEJzpp7p40n.1"
								target="_blank" rel="noopener noreferrer"
								style="display:block;text-decoration:none;color:#ffffff;background:linear-gradient(135deg,#3b82f6,#2563eb);border-radius:12px;padding:12px 20px;font-size:14px;line-height:20px;font-weight:600;font-family:'IBM Plex Sans Thai','Noto Sans Thai','Anuphan','Kanit','Tahoma',Arial,sans-serif;text-align:center;">
								เข้าร่วม Zoom Meeting
							</a>
						</td>
					</tr>
				</table>
			</section>

			<section style="margin-top:32px;">
				<h3
					style="margin:0 0 16px 0;font-size:12px;line-height:16px;font-weight:600;letter-spacing:1.2px;color:#f472b6;text-transform:uppercase;">
					ติดต่อสอบถามเพิ่มเติม
				</h3>

				<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td width="50%" valign="top" style="padding-right:6px;">
							<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"
								style="border:1px solid #fce7f3;background-color:#fdf2f8;border-radius:12px;border-collapse:separate;">
								<tr>
									<td style="padding:14px 16px;">
										<p
											style="margin:0 0 4px 0;font-size:12px;line-height:16px;color:#f472b6;font-weight:500;">
											Line Official Account
										</p>
										<p
											style="margin:0;font-size:14px;line-height:20px;font-weight:600;color:#374151;">
											@triamudom
										</p>
									</td>
								</tr>
							</table>
						</td>

						<td width="50%" valign="top" style="padding-left:6px;">
							<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"
								style="border:1px solid #fce7f3;background-color:#fdf2f8;border-radius:12px;border-collapse:separate;">
								<tr>
									<td style="padding:14px 16px;">
										<p
											style="margin:0 0 4px 0;font-size:12px;line-height:16px;color:#f472b6;font-weight:500;">
											โทรติดต่อ
										</p>
										<p
											style="margin:0;font-size:14px;line-height:20px;font-weight:600;color:#374151;">
											<a href="tel:0803323330"
												style="color:#374151 !important;text-decoration:none !important;font-weight:600;">
												0803323330
											</a>
										</p>
									</td>
								</tr>
							</table>
						</td>
					</tr>
				</table>
			</section>

			<section style="border-top:1px solid #fce7f3;padding-top:24px;margin-top:32px;">
				<p style="margin:0;font-size:14px;line-height:20px;color:#6b7280;font-weight:400;">
					ขอแสดงความนับถือ
				</p>
				<p style="margin:12px 0 0 0;font-size:16px;line-height:24px;font-weight:600;color:#111827;">
					คณะทำงาน Triamudom Family
				</p>
				<p style="margin:2px 0 0 0;font-size:14px;line-height:20px;color:#6b7280;font-weight:400;">
					สมาคมผู้ปกครองและครูโรงเรียนเตรียมอุดมศึกษา
				</p>
			</section>
		</div>

		<div style="border-top:1px solid #fce7f3;background-color:#fdf2f8;padding:16px 32px;">
			<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0">
				<tr>
					<td align="center"
						style="font-size:12px;line-height:16px;color:#f472b6;text-align:center;font-family:'IBM Plex Sans Thai','Noto Sans Thai','Anuphan','Kanit','Tahoma',Arial,sans-serif;font-weight:400;">
						<a href="https://github.com/Triamudom-Family/Staff-EmailAnnouncement?tab=readme-ov-file"
							target="_blank" rel="noopener noreferrer"
							style="color:#f472b6;text-decoration:none;font-weight:400;">
							Github
						</a>

						<span style="color:#fbcfe8;padding:0 6px;">•</span>

						<a href="https://github.com/Triamudom-Family/Staff-EmailAnnouncement?tab=AGPL-3.0-1-ov-file"
							target="_blank" rel="noopener noreferrer"
							style="color:#f472b6;text-decoration:none;font-weight:400;">
							AGPL-3.0
						</a>

						<span style="color:#fbcfe8;padding:0 6px;">|</span>

						<span style="font-weight:400;">Email Send v2.3.2</span>

						<span style="color:#fbcfe8;padding:0 6px;">•</span>

						<span style="font-weight:400;">For Triamudom Family</span>
					</td>
				</tr>
			</table>
		</div>

	</div>
</body>

</html>
	`.trim();
}