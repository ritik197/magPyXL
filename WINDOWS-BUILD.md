# MagpieXL — Windows Par Build Kaise Karein

Is sandbox me Windows-native wheel cross-compile nahi ho paata (Rust ka
Windows target install karne ke liye `rustup` chahiye, jo yahan available
nahi hai). Lekin apne Windows computer pe banana bahut aasan hai — 5 minute
ka kaam, aur genuine Windows toolchain use hoga (cross-compile se zyada
reliable).

## Steps (Windows par)

1. **Rust install karo** (agar pehle se nahi hai):
   https://rustup.rs se `rustup-init.exe` download karo, chalao, default
   options accept karo. Terminal band karke dobara kholo.

   Check karo: `rustc --version`

2. **Python check karo** (3.8+ chahiye): `python --version`

3. **maturin install karo:**
   ```
   pip install maturin
   ```

4. **Source code extract karo** (`magpiexl-source.zip` diya hua hai) — ek
   folder me unzip karo, us folder me terminal khol ke:
   ```
   maturin build --release
   ```

5. Wheel ban jayegi `target\wheels\` folder me (jaise
   `magpiexl-0.1.0-cp312-cp312-win_amd64.whl`). Usse install karo:
   ```
   pip install target\wheels\magpiexl-0.1.0-cp312-cp312-win_amd64.whl
   ```

Bas — ab `import magpiexl` Windows pe bhi kaam karega.

## Aur bhi aasan tarika (agar sirf use karna hai, dobara build nahi)

Step 4-5 ki jagah, seedha project folder me terminal khol ke ye ek command
chala sakte ho — maturin khud build karke install kar dega:
```
pip install .
```

## Note

Ye process sirf ek baar karna padta hai. Baad me wheel file kahin bhi
copy karke normal `pip install <file>.whl` se install ho jayegi — dobara
Rust/maturin ki zarurat nahi padegi us machine pe.
