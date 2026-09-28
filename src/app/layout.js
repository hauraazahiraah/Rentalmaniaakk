import '../styles/globals.css';

export const metadata = {
  title: 'Sampahahayuk',
  description: 'Aplikasi Manajemen Sampahahayuk',
};

export default function RootLayout({ children }) {
  return (
    <html lang="id">
      <body>
        {children}
      </body>
    </html>
  );
}