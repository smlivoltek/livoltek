# LIVOLTEK | Estoque Comercial Brasil

Dashboard comercial de estoque em tempo real — Armazém 07.

## Como publicar no GitHub Pages

### 1. Criar o repositório (só na primeira vez)

1. Acesse [github.com](https://github.com) e clique em **New repository**
2. Nome sugerido: `estoque-comercial` (ou use um repositório já existente)
3. Deixe **Public** (necessário para GitHub Pages gratuito)
4. Clique em **Create repository**

### 2. Fazer upload dos arquivos

1. Dentro do repositório, clique em **Add file → Upload files**
2. Arraste os dois arquivos:
   - `index.html`
   - `README.md`
3. Clique em **Commit changes**

### 3. Ativar o GitHub Pages

1. Vá em **Settings** (engrenagem no topo do repositório)
2. No menu lateral, clique em **Pages**
3. Em **Source**, selecione **Deploy from a branch**
4. Em **Branch**, selecione **main** e pasta **/ (root)**
5. Clique em **Save**
6. Aguarde ~1 minuto e acesse a URL exibida:
   ```
   https://SEU-USUARIO.github.io/estoque-comercial/
   ```

---

## Atualização semanal

Não é necessário alterar o `index.html` para atualizar os dados.

O dashboard busca os dados **diretamente das planilhas Google Sheets publicadas** ao abrir a página. Basta atualizar a planilha e o dashboard refletirá automaticamente ao próximo acesso (ou ao clicar em **Atualizar Dados**).

Se precisar forçar atualização sem reabrir: clique no botão **🔄** no cabeçalho.

---

## Fontes de dados

| Armazém | Aba da planilha | GID |
|---------|----------------|-----|
| CWB — Curitiba | Curitiba | 0 |
| FOR — Fortaleza | Fortaleza | 411796829 |
| MAO — Livoltek | MAO-Livoltek | 2006971885 |
| MAO — Eletra | MAO-Eletra | 1879495818 |
| Reservas Comerciais | Reservas | 577818960 |

Planilha base: `2PACX-1vQXyqYNC-Rv_Hva2AT0Mid_HzzMt1FeqXYnHCiIH8rThS-9wBVMEc4KShgLYSv-JkoIgFh6tCczQjXX`

---

## Lógica de cálculo

```
Disponível Comercial = Saldo TOTVS (col H) − Empenho TOTVS (col I) − Reservas Comerciais
```

- **Saldo TOTVS** — estoque físico registrado no ERP, coluna H da aba Saldos em Estoque
- **Empenho TOTVS** — pedidos já lançados formalmente no ERP (coluna I), desconto automático
- **Reservas Comerciais** — comprometimentos por vendedores ainda não lançados no ERP (aba Reservas), cruzados por Código + Armazém

---

## Estrutura do repositório

```
/
├── index.html    ← dashboard completo (único arquivo necessário)
└── README.md     ← este arquivo
```
