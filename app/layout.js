import "./globals.css";

export const metadata = {
  title: "Offer Letter Generator",
  description: "Generate appointment letters from a Word template"
};

export default function RootLayout({ children }) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  );
}
