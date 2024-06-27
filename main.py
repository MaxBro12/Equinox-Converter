import pandas as pd


def update_new_tables():
    table = pd.ExcelFile('EquinoxSystemResourceData4.xlsx')
    stars = pd.read_excel(table, 'Stars')
    planets = pd.read_excel(table, 'Planets')

    stars = stars.drop(columns=['starID', 'Star'])
    planets = planets.drop(columns=['planetID', 'Planet Name'])

    regions = stars['regionName'].drop_duplicates()

    equinox_table = pd.ExcelWriter('UQUINOX.xlsx')

    static_page = pd.DataFrame(
        columns=['Region', 'Median Power', 'Median Workforce', 'Superionic Ice / Hour', 'Magmatic Gas / Hour']
    )
    for region in regions:
        region_table = pd.DataFrame(
            columns=['System Name', 'power', 'Workforce', 'Superionic Ice / Hour', 'Magmatic Gas / Hour']
        )
        for index, star in stars.loc[stars['regionName'] == region].iterrows():
            star_planets = planets.loc[planets['System Name'] == star['System Name']]
            region_table.loc[len(region_table.index)] = [
                star['System Name'],
                star['power'] + star_planets['Power'].sum(),
                star_planets['Workforce'].sum(),
                star_planets['Superionic Ice / Hour'].sum(),
                star_planets['Magmatic Gas / Hour'].sum(),
            ]

        region_table.to_excel(equinox_table, sheet_name=region, index=False)

        static_page.loc[len(static_page.index)] = [
            region,
            region_table['power'].mean().round(0),
            region_table['Workforce'].mean().round(0),
            region_table['Superionic Ice / Hour'].sum(),
            region_table['Magmatic Gas / Hour'].sum(),
        ]

    static_page.to_excel(equinox_table, sheet_name='STATISTIC', index=False)
    equinox_table.close()


if __name__ == '__main__':
    update_new_tables()

