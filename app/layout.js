export const metadata = {
  title: 'SNV Slide Generator',
  description: 'Generate PowerPoint slides with ease',
};

export default function RootLayout({ children }) {
  return (
    <html lang="en">
      <body style={{ margin: 0, fontFamily: '"Tahoma","MS Sans Serif",Arial,sans-serif', background: '#008080' }}>
        {children}
      </body>
    </html>
  );
}
