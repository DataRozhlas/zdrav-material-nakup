# %%
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
import matplotlib as mpl

# %%
d = pd.concat(pd.read_excel('./data/Přehled-objednávek-k-18.-5.-2020.xlsx', sheet_name=None, header=None, names=['datum', 'partner', 'kategorie', 'mnozstvi', 'kus_bez_dph', 'celkem_bez_dph']), ignore_index=True)

# %%
d.fillna(value='', inplace=True)
d = d[d.celkem_bez_dph.str.contains('Kč')]

# %%
d.groupby(['partner', 'kategorie']).celkem_bez_dph.sum().to_excel('./agregace.xlsx')

# %%
def fix(val):
    return float(str(val).replace(' ', '').replace('Kč', '').replace(',', '.'))
d[['mnozstvi', 'kus_bez_dph', 'celkem_bez_dph']] = d[['mnozstvi', 'kus_bez_dph', 'celkem_bez_dph']].applymap(fix)

# %%
def cln(val):
    return val.replace('\n', ' ')

d[['partner', 'kategorie']] = d[['partner', 'kategorie']].applymap(cln)

# %%
# filtr
#d = d[d.kategorie.str.contains('ÚSTENKY')]

# %%
celk = pd.DataFrame(d.groupby(['partner']).celkem_bez_dph.sum())
celk.reset_index(inplace=True)
celk = celk.sort_values('celkem_bez_dph', ascending=True).head(20)

celk.celkem_bez_dph = celk.celkem_bez_dph / 1000000

# %%
d.groupby(['partner']).celkem_bez_dph.sum().to_excel('./agg_celk.xlsx')

# %%
d[d.kategorie.str.contains('FFP2')].groupby(['partner']).celkem_bez_dph.sum().to_excel('./agg_ffp2.xlsx')

# %%
d[d.kategorie.str.contains('FFP3')].groupby(['partner']).celkem_bez_dph.sum().to_excel('./agg_ffp3.xlsx')

# %%
plt.style.use('./example.mplstyle')
fig, ax = plt.subplots(figsize=(8,6))
ax.barh(celk.partner, celk.celkem_bez_dph)
ax.xaxis.grid(True)
ax.spines['top'].set_visible(False)
ax.spines['bottom'].set_visible(False)
ax.spines['right'].set_visible(False)
ax.xaxis.set_ticks_position('none')
ax.yaxis.set_ticks_position('none') 

ax.set_xlabel('miliony Kč, bez DPH')
ax.set_title('Nákupy ústenek')
plt.tight_layout()
plt.savefig('tst.svg')

# %%
