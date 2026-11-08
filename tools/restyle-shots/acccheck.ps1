param(
    [Parameter(Mandatory = $true)][string]$In,
    [string]$Combo = "skyl",
    [int]$Stride = 3,
    [double]$Tolerance = 26,
    [double]$MinShare = 0.0008,
    [int]$Top = 12
)

# Palette audit for one screenshot: histogram the pixels and compare every frequent color
# against the Theme token table of that skin/mode combo. A color counts as "on palette" if
# it is close to a token itself or to the segment between two tokens, because anti-aliased
# and ClearType pixels are always blends of two token colors. Frequent colors farther from
# every segment than Tolerance are suspected hardcoded colors that bypassed the token layer.
# KNOWN BLIND SPOT: because white is a token (text colour), any light grey also sits close to
# the white-to-muted segment, so a stale native fill (e.g. the TabControl strip at 236,236,236
# in a dark window) passes this check. Large flat areas still need the eyeball pass on the crop.
$Palettes = @{
    skyl  = @('F1F7FC','E7F2FA','FFFFFF','C3DCEF','DCEBF7','EAF4FC','143C64','5A82AA','2E8BE0','D8ECFB','E3EFF9','F7FBFE','E2EFF8','EFF6FC','1565C0','E07B39','C43C31','000000')
    skyd  = @('161A20','1C222A','1F2630','2E3947','263041','232C38','E3EEF9','92A6BC','4E9BE8','23384F','2A3543','1D232C','28313C','7FC0F5','F0A070','FF6B6B','FFFFFF','000000')
    densl = @('F2F3F5','E9EBEF','FFFFFF','D2D6DC','E4E7EB','EEF0F3','1B1D21','5F666E','3A6EA5','D9E6F5','DDE1E6','F8F9FB','E7E9ED','2A5FA8','B4442B','B03A2E','000000')
    dense = @('1D1F22','23262A','212428','35393F','2B2F34','262A2F','DCE0E5','8B929B','5B8DD0','2C3F5A','2F343A','232629','2E3238','7FB0E8','E08366','F0796B','FFFFFF','000000')
}

$hexes = $Palettes[$Combo]
if ($null -eq $hexes) { Write-Output "unknown combo=$Combo"; exit 1 }
# ClearType overshoots into saturated R/G/B at glyph edges; those fringes are not
# hardcoded colors, so allow the three primaries and three secondaries as blend endpoints too.
$hexes = $hexes + @('FF0000','00FF00','0000FF','FFFF00','00FFFF','FF00FF')
$rgb = foreach ($h in $hexes) { [Convert]::ToInt32($h, 16) }

Add-Type -AssemblyName System.Drawing
Add-Type -ReferencedAssemblies @('System.Drawing') -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

public static class PalAudit {
    public static string Run(string path, int[] pal, int stride, double tol, double minShare, int top) {
        var hist = new Dictionary<int, long>();
        long sampled = 0;
        int imgW = 0, imgH = 0;
        byte[] full = null;
        int fullStride = 0;
        using (var bmp = new Bitmap(path)) {
            int w = bmp.Width, h = bmp.Height;
            imgW = w; imgH = h;
            var rect = new Rectangle(0, 0, w, h);
            BitmapData data = bmp.LockBits(rect, ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
            byte[] buf = new byte[data.Stride * h];
            Marshal.Copy(data.Scan0, buf, 0, buf.Length);
            bmp.UnlockBits(data);
            full = buf; fullStride = data.Stride;
            for (int y = 0; y < h; y++) {
                int row = y * data.Stride;
                for (int x = 0; x < w; x += stride) {
                    int i = row + x * 4;
                    int k = ((buf[i + 2] & 0xFF) << 16) | ((buf[i + 1] & 0xFF) << 8) | (buf[i] & 0xFF);
                    long c;
                    hist.TryGetValue(k, out c);
                    hist[k] = c + 1;
                    sampled++;
                }
            }
        }
        if (sampled == 0) return "sampled=0 FAIL";
        var off = new List<string>();
        var fringe = new List<string>();
        foreach (var kv in hist) {
            double share = 100.0 * kv.Value / sampled;
            if (share / 100.0 < minShare) continue;
            double d = Dist(kv.Key, pal);
            if (d > tol) {
                // ClearType paints each glyph edge with one channel pushed past the text colour,
                // so a light-on-dark window always shows a few frequent blue/orange pixel values.
                // A real hardcoded widget colour fills an area, so it survives as a long horizontal
                // run; a glyph fringe does not. Only report colours that actually paint something.
                int run = MaxRun(full, fullStride, imgW, imgH, kv.Key);
                if (run >= 8) off.Add(string.Format("OFF {0:X6} share={1:0.000}% dist={2:0.0} maxRun={3}", kv.Key, share, d, run));
                else fringe.Add(string.Format("fringe {0:X6} share={1:0.000}% dist={2:0.0} maxRun={3}", kv.Key, share, d, run));
            }
        }
        off.Sort();
        var sb = new StringBuilder();
        sb.AppendFormat("size={0} sampled={1} distinct={2}", Path.GetFileName(path), sampled, hist.Count);
        int n = Math.Min(top, off.Count);
        for (int i = 0; i < n; i++) sb.AppendLine().Append(off[i]);
        if (off.Count == 0 && fringe.Count > 0)
            sb.Append("  (cleartype fringe skipped: ").Append(fringe.Count).Append(')');
        sb.AppendFormat("\nOFFENDERS={0} {1}", off.Count, off.Count == 0 ? "PASS" : "FAIL");
        return sb.ToString();
    }

    // Longest horizontal run of pixels within a small distance of the target colour, full resolution.
    static int MaxRun(byte[] buf, int stride, int w, int h, int rgb) {
        int tr = (rgb >> 16) & 0xFF, tg = (rgb >> 8) & 0xFF, tb = rgb & 0xFF;
        int best = 0;
        for (int y = 0; y < h; y++) {
            int row = y * stride;
            int run = 0;
            for (int x = 0; x < w; x++) {
                int i = row + x * 4;
                int dr = (buf[i + 2] & 0xFF) - tr, dg = (buf[i + 1] & 0xFF) - tg, db = (buf[i] & 0xFF) - tb;
                if (dr * dr + dg * dg + db * db <= 27) {
                    run++;
                    if (run > best) best = run;
                } else run = 0;
            }
            if (best >= 8) return best;
        }
        return best;
    }

    static double Dist(int rgb, int[] pal) {
        int r = (rgb >> 16) & 0xFF, g = (rgb >> 8) & 0xFF, b = rgb & 0xFF;
        double best = 1e9;
        for (int i = 0; i < pal.Length; i++) {
            int ai = pal[i];
            int ar = (ai >> 16) & 0xFF, ag = (ai >> 8) & 0xFF, ab = ai & 0xFF;
            for (int j = i; j < pal.Length; j++) {
                int bi2 = pal[j];
                int br = (bi2 >> 16) & 0xFF, bg = (bi2 >> 8) & 0xFF, bb = bi2 & 0xFF;
                double dr = br - ar, dg = bg - ag, db = bb - ab;
                double len2 = dr * dr + dg * dg + db * db;
                double t = 0;
                if (len2 > 0) {
                    t = ((r - ar) * dr + (g - ag) * dg + (b - ab) * db) / len2;
                    if (t < 0) t = 0; else if (t > 1) t = 1;
                }
                double er = r - (ar + t * dr), eg = g - (ag + t * dg), eb = b - (ab + t * db);
                double d = Math.Sqrt(er * er + eg * eg + eb * eb);
                if (d < best) best = d;
            }
        }
        return best;
    }
}
'@

$out = [PalAudit]::Run($In, $rgb, [Math]::Max(1, $Stride), $Tolerance, $MinShare, $Top)
Write-Output "image=$In combo=$Combo"
Write-Output $out
