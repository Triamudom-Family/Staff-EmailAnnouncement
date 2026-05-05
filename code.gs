function myFunction() {
	Logger.log(MailApp.getRemainingDailyQuota());

	for (let i = 2; i <= 25; i++) {
		sendEmailByRow(i);
	}
}

function sendEmailByRow(row) {
	const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Finalists");
	if (!sheet) throw new Error('Sheet not found');

	row = Number(row);
	if (!row || row < 2) throw new Error('Invalid row number');

	const statusCol = 15;

	try {
		if (MailApp.getRemainingDailyQuota() < 1) throw new Error('No email quota remaining today');

		const values = sheet.getRange(row, 2, 1, 13).getValues()[0];

		const emailAddress = String(values[0] || "").trim();
		const teamName = String(values[1] || "").trim();
		const memberCount = Number(values[2]) || "";

		if (!emailAddress) throw new Error("Missing email address in column B");
		if (!teamName) throw new Error("Missing team name in column C");

		const members = [];

		for (let i = 0; i < 5; i++) {
			const firstName = String(values[3 + i * 2] || "").trim();
			const lastName = String(values[4 + i * 2] || "").trim();

			if (firstName || lastName) {
				members.push({
					firstName,
					lastName,
					fullName: `${firstName} ${lastName}`.trim()
				});
			}
		}

		const emailData = {
			emailAddress,
			teamName,
			memberCount,
			members,

			member1FirstName: members[0]?.firstName || "",
			member1LastName: members[0]?.lastName || "",
			member2FirstName: members[1]?.firstName || "",
			member2LastName: members[1]?.lastName || "",
			member3FirstName: members[2]?.firstName || "",
			member3LastName: members[2]?.lastName || "",
			member4FirstName: members[3]?.firstName || "",
			member4LastName: members[3]?.lastName || "",
			member5FirstName: members[4]?.firstName || "",
			member5LastName: members[4]?.lastName || ""
		};

		const subject = "ประกาศผลการคัดเลือกเข้าโครงการ OASIS";
		const plainText = generatePlainText(emailData);
		const htmlBody = generateHtmlBody(emailData);

		GmailApp.sendEmail(emailAddress, subject, plainText, {
			htmlBody: htmlBody,
			name: "Organtogo Gen 2",
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
	const memberLines = (data.members || [])
		.map((member, index) => `${index + 1}. ${member.fullName}`)
		.join("\n");

	return `
ประกาศผลการคัดเลือก

ยินดีด้วยกับทีม
${data.teamName}

ผ่านการคัดเลือกเข้าโครงการ Organ Ambassador Student Innovation Scheme: OASIS

สมาชิกในทีม (${data.members.length} คน)
${memberLines}

ทาง Organtogo Gen 2 ในเครือข่ายเยาวชนอาสาสมัครกาชาดไทย (Thai Red Cross Youth Network: TRCYN)
ภายใต้สำนักยุวกาชาดและอาสาสมัคร สภากาชาดไทย ขอแสดงความยินดีเป็นอย่างยิ่ง
เนื่องในโอกาสที่ท่านได้รับการคัดเลือกให้เข้าร่วมเป็นยุวทูตอวัยวะ

การปฐมนิเทศ
วันที่: 11 พฤษภาคม พ.ศ. 2569
สถานที่: อาคารรัตนวิทยาพัฒน์ โรงพยาบาลจุฬาลงกรณ์ สภากาชาดไทย

การยืนยันสิทธิ์เข้าร่วมโครงการ
กรุณาดำเนินการยืนยันสิทธิ์การเข้าร่วมโครงการภายในวันที่ 5 พฤษภาคม พ.ศ. 2569
ผ่านแบบฟอร์มด้านล่างนี้

https://docs.google.com/forms/d/e/1FAIpQLSe1lzaPweqw__nyjQLErsl6gXsZUxf01JUyslbd1vlMI1YIJQ/viewform?usp=dialog

หากท่านไม่ดำเนินการยืนยันสิทธิ์ภายในระยะเวลาที่กำหนด
ทางโครงการขอสงวนสิทธิ์ในการพิจารณาว่าท่านสละสิทธิ์ในการเข้าร่วมโครงการ

ติดต่อสอบถามเพิ่มเติม
Instagram: @organtogoth
Email: organtogo@gmail.com

ขอแสดงความนับถือ

นางสาวบุลภรณ์ รัตนะเจริญธรรม
Head of Human Resources
	`.trim();
}

function generateHtmlBody(data) {
	const memberRows = (data.members || []).map((member, index) => {
		const border = index === data.members.length - 1 ? "" : "border-bottom:1px solid #f3f4f6;";

		return `
		<tr>
			<td style="padding:12px 16px;${border}">
				<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td width="28" valign="middle" style="width:28px;">
							<div style="width:24px;height:24px;border-radius:50%;background-color:#ffe4e6;color:#e11d48;font-size:11px;font-weight:500;line-height:24px;text-align:center;">
								${index + 1}
							</div>
						</td>
						<td valign="middle" style="padding-left:12px;font-size:14px;line-height:20px;color:#374151;font-weight:400;">
							${escapeHtml_(member.fullName)}
						</td>
					</tr>
				</table>
			</td>
		</tr>
	`;
	}).join("");

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
</head>

<body
	style="margin:0; box-sizing:border-box; width:100%; -webkit-text-size-adjust:100%; font-family:'IBM Plex Sans Thai','Noto Sans Thai', 'Anuphan', 'Kanit', 'Noto Sans Thai','Tahoma', Arial, sans-serif; background-image:linear-gradient(to bottom, #fff1f2, #ffffff, #f3f4f6); min-height:100vh; padding-left:1.5rem; padding-right:1.5rem; padding-top:2.5rem; padding-bottom:2.5rem; color:#1f2937;">
	<div
		style="margin-left:auto; margin-right:auto; max-width:48rem; overflow:hidden; border-radius:1.5rem; border-style:solid; border-width:1px; border-color:#e5e7eb; background-color:#ffffff; box-shadow:0 20px 25px -5px rgba(0,0,0,0.10), 0 8px 10px -6px rgba(0,0,0,0.10); width:100%; box-sizing:border-box; font-family:'IBM Plex Sans Thai','Noto Sans Thai', 'Anuphan', 'Kanit', 'Noto Sans Thai','Tahoma', Arial, sans-serif;">

		<div
			style="background-image:linear-gradient(to top right, #fb7185, #ef4444, #dc2626); padding-left:3rem; padding-right:3rem; padding-top:2.5rem; padding-bottom:2.5rem; color:#ffffff; text-align:center; position:relative;">
			<div
				style="margin-bottom:1rem; display:inline-flex; align-items:center; gap:0.5rem; border-radius:9999px; background-color:rgba(255,255,255,0.15); border-style:solid; border-width:1px; border-color:rgba(255,255,255,0.30); padding-left:1rem; padding-right:1rem; padding-top:0.25rem; padding-bottom:0.25rem; font-size:0.75rem; line-height:1rem;">
				ประกาศผลการคัดเลือก
			</div>
			<h1 style="margin:0; font-size:2.25rem; line-height:2.5rem; font-weight:700; letter-spacing:-0.025em;">
				ยินดีด้วยกับทีม</h1>
			<h2
				style="margin:0; margin-top:0.5rem; font-size:1.875rem; line-height:2.25rem; font-weight:600; opacity:0.9;">
				${escapeHtml_(data.teamName)}</h2>
			<div
				style="margin-top:1.25rem; display:inline-flex; align-items:center; gap:0.5rem; border-radius:9999px; background-color:rgba(255,255,255,0.15); border-style:solid; border-width:1px; border-color:rgba(255,255,255,0.30); padding-left:1rem; padding-right:1rem; padding-top:0.5rem; padding-bottom:0.5rem; font-size:0.875rem; line-height:1.25rem;">
				ผ่านการคัดเลือกเข้าโครงการ Organ Ambassador Student Innovation Scheme: OASIS
			</div>
		</div>

		<div style="padding-left:3rem; padding-right:3rem; padding-top:2rem; padding-bottom:2rem;">

			<section>
				<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"
					style="margin-bottom:16px;">
					<tr>
						<td align="left" valign="middle">
							<h3
								style="margin:0;font-size:12px;line-height:16px;font-weight:600;letter-spacing:1.2px;color:#9ca3af;text-transform:uppercase;">
								สมาชิกในทีม
							</h3>
						</td>

						<td align="right" valign="middle">
							<span
								style="display:inline-block;border-radius:9999px;background-color:#fff1f2;border:1px solid #ffe4e6;padding:4px 12px;font-size:12px;line-height:16px;font-weight:500;color:#f43f5e;white-space:nowrap;">
								${data.members.length} คน
							</span>
						</td>
					</tr>
				</table>
				<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"
					style="border:1px solid #f3f4f6;border-radius:16px;border-collapse:separate;overflow:hidden;">
					${memberRows}
				</table>
			</section>

			<section style="margin-top:2rem;">
				<p style="margin:0; font-size:1rem; line-height:1.75rem; color:#374151;">
					ทาง Organtogo Gen 2 ในเครือข่ายเยาวชนอาสาสมัครกาชาดไทย (Thai Red Cross Youth Network: TRCYN)
					ภายใต้สำนักยุวกาชาดและอาสาสมัคร สภากาชาดไทย ขอแสดงความยินดีเป็นอย่างยิ่ง
					เนื่องในโอกาสที่ท่านได้รับการคัดเลือกให้เข้าร่วมเป็น<span
						style="font-weight:600; color:#111827;">ยุวทูตอวัยวะ</span>
				</p>
			</section>

			<section style="margin-top:2rem;">
				<h3
					style="margin:0; font-size:0.75rem; line-height:1rem; font-weight:600; letter-spacing:0.1em; color:#9ca3af; text-transform:uppercase; margin-bottom:1rem;">
					การปฐมนิเทศ</h3>
				<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td width="50%" valign="top" style="padding-right:6px;">
							<table role="presentation" width="100%" height="110" cellpadding="0" cellspacing="0"
								border="0"
								style="height:100px;border:1px solid #ffe4e6;background-color:#fff1f2;border-radius:16px;border-collapse:separate;">
								<tr>
									<td valign="top" style="padding:20px;">
										<p
											style="margin:0 0 4px 0;font-size:12px;line-height:16px;color:#fb7185;font-weight:500;">
											วันที่
										</p>
										<p
											style="margin:0;font-size:14px;line-height:20px;font-weight:600;color:#9f1239;">
											11 พฤษภาคม พ.ศ. 2569
										</p>
									</td>
								</tr>
							</table>
						</td>

						<td width="50%" valign="top" style="padding-left:6px;">
							<table role="presentation" width="100%" height="110" cellpadding="0" cellspacing="0"
								border="0"
								style="height:100px;border:1px solid #ffe4e6;background-color:#fff1f2;border-radius:16px;border-collapse:separate;">
								<tr>
									<td valign="top" style="padding:20px;">
										<p
											style="margin:0 0 4px 0;font-size:12px;line-height:16px;color:#fb7185;font-weight:500;">
											สถานที่
										</p>
										<p
											style="margin:0;font-size:14px;line-height:20px;font-weight:600;color:#9f1239;">
											อาคารรัตนวิทยาพัฒน์ โรงพยาบาลจุฬาลงกรณ์ สภากาชาดไทย
										</p>
									</td>
								</tr>
							</table>
						</td>
					</tr>
				</table>
			</section>

			<section style="margin-top:2rem;">
				<h3
					style="margin:0; font-size:0.75rem; line-height:1rem; font-weight:600; letter-spacing:0.1em; color:#9ca3af; text-transform:uppercase; margin-bottom:1rem;">
					การยืนยันสิทธิ์เข้าร่วมโครงการ</h3>
				<p style="margin:0; font-size:1rem; line-height:1.75rem; color:#374151;">
					กรุณาดำเนินการยืนยันสิทธิ์การเข้าร่วมโครงการภายในวันที่
					<span style="font-weight:600; color:#111827;">5 พฤษภาคม พ.ศ. 2569</span>
					ผ่านแบบฟอร์มด้านล่างนี้
				</p>
				<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"
					style="margin-top:1.25rem;">
					<tr>
						<td align="center" style="text-align:center;">
							<a href="https://docs.google.com/forms/d/e/1FAIpQLSe1lzaPweqw__nyjQLErsl6gXsZUxf01JUyslbd1vlMI1YIJQ/viewform?usp=dialog"
								target="_blank" rel="noopener noreferrer"
								style="display:inline-block;text-decoration:none;color:#ffffff;background-color:#ef4444;border-radius:12px;padding:12px 20px;font-size:14px;line-height:20px;font-weight:600;font-family:'IBM Plex Sans Thai','Noto Sans Thai','Anuphan','Kanit','Noto Sans Thai','Tahoma', Arial,sans-serif;">
								กรอกแบบฟอร์มยืนยันสิทธิ์การเข้าร่วมโครงการ
							</a>
						</td>
					</tr>
				</table>
				<div
					style="margin-top:1rem; border-radius:0.75rem; border-style:solid; border-width:1px; border-color:#e5e7eb; background-color:#fafafa; padding-left:1rem; padding-right:1rem; padding-top:1rem; padding-bottom:1rem; font-size:0.875rem; line-height:1.75rem; color:#374151;">
					หากท่านไม่ดำเนินการยืนยันสิทธิ์ภายในระยะเวลาที่กำหนด
					ทางโครงการขอสงวนสิทธิ์ในการพิจารณาว่าท่าน<span
						style="font-weight:600; color:#dc2626;">สละสิทธิ์</span>ในการเข้าร่วมโครงการ
				</div>
			</section>

			<section style="margin-top:32px;">
				<h3
					style="margin:0 0 16px 0;font-size:12px;line-height:16px;font-weight:600;letter-spacing:1.2px;color:#9ca3af;text-transform:uppercase;">
					ติดต่อสอบถามเพิ่มเติม
				</h3>

				<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td width="50%" valign="top" style="padding-right:6px;">
							<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"
								style="border:1px solid #f3f4f6;background-color:#f9fafb;border-radius:12px;border-collapse:separate;">
								<tr>
									<td style="padding:14px 16px;">
										<p
											style="margin:0 0 4px 0;font-size:12px;line-height:16px;color:#9ca3af;font-weight:500;">
											Instagram
										</p>
										<p
											style="margin:0;font-size:14px;line-height:20px;font-weight:600;color:#374151;">
											@organtogoth
										</p>
									</td>
								</tr>
							</table>
						</td>

						<td width="50%" valign="top" style="padding-left:6px;">
							<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"
								style="border:1px solid #f3f4f6;background-color:#f9fafb;border-radius:12px;border-collapse:separate;">
								<tr>
									<td style="padding:14px 16px;">
										<p
											style="margin:0 0 4px 0;font-size:12px;line-height:16px;color:#9ca3af;font-weight:500;">
											Email
										</p>
										<p style="margin:0;font-size:14px;line-height:20px;font-weight:600;">
											<a href="mailto:organtogo@gmail.com"
												style="color:#374151 !important;text-decoration:none !important;font-weight:600;">
												organtogo@gmail.com
											</a>
										</p>
									</td>
								</tr>
							</table>
						</td>
					</tr>
				</table>
			</section>
			<section
				style="border-top-style:solid; border-top-width:1px; border-color:#f3f4f6; padding-top:0.25rem; margin-top:2rem;">
				<p style="margin:0; font-size:0.875rem; line-height:1.25rem; color:#6b7280;">ขอแสดงความนับถือ</p>
				<p
					style="margin:0; margin-top:0.75rem; font-size:1rem; line-height:1.5rem; font-weight:600; color:#111827;">
					นางสาวบุลภรณ์ รัตนะเจริญธรรม</p>
				<p style="margin:0; margin-top:0.125rem; font-size:0.875rem; line-height:1.25rem; color:#6b7280;">Head
					of Human Resources</p>
			</section>

		</div>
		<div style="border-top:1px solid #f3f4f6;background-color:#f9fafb;padding:16px 32px;">
			<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0">
				<tr>
					<td align="center"
						style="font-size:12px;line-height:16px;color:#9ca3af;text-align:center;font-family:'IBM Plex Sans Thai','Noto Sans Thai','Anuphan','Kanit','Noto Sans Thai','Tahoma',Arial,sans-serif;">

						<a href="https://github.com/arckanop/OASIS-Mail-Announcement" target="_blank"
							rel="noopener noreferrer" style="color:#9ca3af;text-decoration:none;font-weight:400;">
							Github
						</a>

						<span style="color:#d1d5db;padding:0 6px;">•</span>

						<a href="https://github.com/arckanop/OASIS-Mail-Announcement/blob/master/LICENSE.md"
							target="_blank" rel="noopener noreferrer"
							style="color:#9ca3af;text-decoration:none;font-weight:400;">
							AGPL-3.0
						</a>

						<span style="color:#d1d5db;padding:0 6px;">|</span>

						<span style="font-weight:400;">Email Send v1.1.2</span>

						<span style="color:#d1d5db;padding:0 6px;">•</span>

						<span style="font-weight:400;">For Organtogo</span>

					</td>
				</tr>
			</table>
		</div>

	</div>
</body>

</html>
	`.trim();
}

