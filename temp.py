import geopandas as gpd
df = gpd.read_file("zip:///workspaces/territory-analysis/data/Milwaukee_Datapoints07072026.zip")
print(df.columns)