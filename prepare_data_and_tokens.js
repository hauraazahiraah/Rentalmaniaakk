const { PrismaClient } = require('@prisma/client');
const jwt = require('jsonwebtoken');
const bcrypt = require('bcryptjs');
const fs = require('fs');

const prisma = new PrismaClient();
const JWT_SECRET = process.env.JWT_SECRET || 'sampahahayuk-secret-key-2024';

async function main() {
  const hashedPassword = await bcrypt.hash('admin123', 10);
  
  // Ensure Admin exists
  const admin = await prisma.user.upsert({
    where: { email: 'admin@sampahahayuk.com' },
    update: {},
    create: {
      name: 'Admin Sampahahayuk',
      email: 'admin@sampahahayuk.com',
      password: hashedPassword,
      phone: '081234567890',
      role: 'ADMIN'
    }
  });

  // Ensure regular User exists
  const user = await prisma.user.upsert({
    where: { email: 'budi@sampahahayuk.com' },
    update: {},
    create: {
      name: 'Budi Santoso',
      email: 'budi@sampahahayuk.com',
      password: hashedPassword,
      phone: '081298765432',
      role: 'USER',
      saldo: 175000
    }
  });

  // Ensure some sample reports exist so user & admin dashboards look awesome!
  const jenisPlastik = await prisma.jenisSampah.findFirst({ where: { name: 'Botol Plastik' } });
  const jenisKardus = await prisma.jenisSampah.findFirst({ where: { name: 'Kardus' } });
  const jenisLogam = await prisma.jenisSampah.findFirst({ where: { name: 'Logam' } });
  const wilayahGambir = await prisma.wilayah.findFirst({ where: { name: 'Gambir' } });
  const wilayahMenteng = await prisma.wilayah.findFirst({ where: { name: 'Menteng' } });

  if (jenisPlastik && wilayahGambir) {
    const existingLaporan = await prisma.laporanSampah.findFirst({ where: { userId: user.id } });
    if (!existingLaporan) {
      await prisma.laporanSampah.create({
        data: {
          userId: user.id,
          jenisSampahId: jenisPlastik.id,
          wilayahId: wilayahGambir.id,
          weight: 15.5,
          totalPrice: 31000,
          status: 'APPROVED',
          tanggalLaporan: new Date(),
        }
      });

      if (jenisKardus && wilayahMenteng) {
        await prisma.laporanSampah.create({
          data: {
            userId: user.id,
            jenisSampahId: jenisKardus.id,
            wilayahId: wilayahMenteng.id,
            weight: 24.0,
            totalPrice: 43200,
            status: 'PENDING',
            tanggalLaporan: new Date(Date.now() - 86400000),
          }
        });
      }
    }
  }

  const adminToken = jwt.sign(
    { id: admin.id, email: admin.email, role: admin.role },
    JWT_SECRET,
    { expiresIn: '7d' }
  );

  const userToken = jwt.sign(
    { id: user.id, email: user.email, role: user.role },
    JWT_SECRET,
    { expiresIn: '7d' }
  );

  fs.writeFileSync('auth_tokens.json', JSON.stringify({
    adminToken,
    userToken,
    adminEmail: admin.email,
    userEmail: user.email
  }, null, 2));

  console.log('Successfully prepared data and generated auth tokens!');
}

main()
  .catch(console.error)
  .finally(() => prisma.$disconnect());
