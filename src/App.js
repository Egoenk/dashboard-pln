import { useState } from "react";
import { Routes, Route, useLocation } from "react-router-dom";
import Topbar from "./scenes/global/Topbar";
import Sidebar from "./scenes/global/Sidebar";
import Dashboard from "./scenes/dashboard";
import Team from "./scenes/team";
import Invoices from "./scenes/invoices";
import Contacts from "./scenes/contacts";
import Bar from "./scenes/bar";
import Form from "./scenes/form";
import Line from "./scenes/line";
import Pie from "./scenes/pie";
import FAQ from "./scenes/faq";
import Geography from "./scenes/geography";
import { CssBaseline, ThemeProvider } from "@mui/material";
import { ColorModeContext, useMode } from "./theme";
import Calendar from "./scenes/calendar/calendar";
import SignInPage from "./scenes/authentication";
import DetailedPage from "./components/DetailedPage/DetailedPage";
import Penjualan from "../src/scenes/Data/Penjualan/index"
import Pendapatan from "../src/scenes/Data/Pendapatan/index"
import Rating from "../src/scenes/Data/Rating/index";
import PLNMobile from "../src/scenes/Data/PLNMobile/index"
import RPTRCT from "../src/scenes/Data/RPTRCT/index"
import HPL from "../src/scenes/Data/HPL/index"
import Pelanggan from "../src/scenes/Data/PenyambunganPelanggan"
import Retur from "../src/scenes/Data/Retur/index"
import SaldoTunggakan from "../src/scenes/Data/SaldoTunggakan/index"
import PiutangPrabayar from "../src/scenes/Data/PiutangPrabayar/index"
import { xlsxToJson } from './path/to/xlsxUtils';


function App() {
  const [theme, colorMode] = useMode();
  const [isSidebar, setIsSidebar] = useState(true);
  const location = useLocation();


  // Cek apakah di halaman login
  const isLoginRoute = location.pathname === '/';

  return (
    <ColorModeContext.Provider value={colorMode}>
      <ThemeProvider theme={theme}>
        <CssBaseline />
        <div className="app">
        {!isLoginRoute && <Sidebar isSidebar={isSidebar} />} {/* Kondisional rendering */}
          <main className={`content ${isLoginRoute ? 'full-width' : ''}`}>
            {!isLoginRoute && <Topbar setIsSidebar={setIsSidebar} />} {/* Kondisional rendering */}
            <Routes>
              <Route path="/" element={<SignInPage/>}/>
              <Route path="/team" element={<Team />} />
              <Route path="/contacts" element={<Contacts />} />
              <Route path="/invoices" element={<Invoices />} />
              <Route path="/form" element={<Form />} />
              <Route path="/bar" element={<Bar />} />
              <Route path="/pie" element={<Pie />} />
              <Route path="/line" element={<Line />} />
              <Route path="/faq" element={<FAQ />} />
              <Route path="/calendar" element={<Calendar />} />
              <Route path="/geography" element={<Geography />} />
              <Route path="/dashboard" element={<Dashboard />} />
              <Route path="/detailedpage" element={<DetailedPage />} />
              <Route path="/penjualan" element={<Penjualan />} />
              <Route path="/pendapatan" element={<Pendapatan />} />
              <Route path="/rating" element={<Rating />} />
              <Route path="/plnmobile" element={<PLNMobile />} />
              <Route path="/rptrct" element={<RPTRCT />} />
              <Route path="/hpl" element={<HPL />} />
              <Route path="/penyambunganpelanggan" element={<Pelanggan />} />
              <Route path="/retur" element={<Retur />} />
              <Route path="/saldotunggakan" element={<SaldoTunggakan />} />
              <Route path="/piutangprabayar" element={<PiutangPrabayar />} />
            </Routes>
          </main>
        </div>
      </ThemeProvider>
    </ColorModeContext.Provider>
  );
}

export default App;
