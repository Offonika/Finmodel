import pandas as pd


def aggregate_data(df: pd.DataFrame) -> pd.DataFrame:
    """Aggregate numeric values by organization.

    If the input DataFrame contains a column called ``'organization'``, all
    numeric columns are summed for each organization.  The resulting DataFrame
    is returned with ``organization`` preserved.  If such a column is missing
    the DataFrame is returned unchanged.
    """
    if 'organization' not in df.columns:
        return df

    numeric_cols = df.select_dtypes(include='number').columns
    aggregated = df.groupby('organization')[list(numeric_cols)].sum().reset_index()
    # Keep non-numeric columns that are constant per group
    other_cols = [c for c in df.columns if c not in numeric_cols and c != 'organization']
    for col in other_cols:
        first_values = df.groupby('organization')[col].first()
        aggregated[col] = aggregated['organization'].map(first_values)
    # Reorder columns to match original layout
    cols_order = ['organization'] + [c for c in df.columns if c != 'organization']
    aggregated = aggregated[[c for c in cols_order if c in aggregated.columns]]
    return aggregated
