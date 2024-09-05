from datetime import datetime
import pandas
import matplotlib.pyplot as plt

data = pandas.read_excel('data.xlsx')
data['receiving_date'] = pandas.to_datetime(data['receiving_date'], format='%d.%m.%Y', errors='coerce')
reference_date = datetime(2021, 7, 1)


def question1():
    july_data = data[(data['receiving_date'].dt.month == 7) & (data['receiving_date'].dt.year == 2021) & (
            data['status'] != 'ПРОСРОЧЕНО')]
    total_revenue_july = july_data['sum'].sum()
    print(f'Общая выручка за июль 2021: {total_revenue_july}')


def question2():
    data['month_year'] = data['receiving_date'].dt.to_period('M')
    monthly_revenue = data.groupby('month_year')['sum'].sum()
    monthly_revenue.plot(kind='bar')
    plt.title('Выручка компании по месяцам')
    plt.xlabel('Месяц')
    plt.ylabel('Выручка')
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.show()


def question3():
    september_data = data[(data['receiving_date'].dt.month == 9) & (data['receiving_date'].dt.year == 2021)]
    manager_revenue = september_data.groupby('sale')['sum'].sum()
    top_manager = manager_revenue.idxmax()
    print(f'Менеджер, привлекший больше всего денежных средств в сентябре 2021: {top_manager}')


def question4():
    october_data = data[(data['receiving_date'].dt.month == 10) & (data['receiving_date'].dt.year == 2021)]
    deal_types = october_data['new/current'].value_counts()
    dominant_deal_type = deal_types.idxmax()
    print(f'Преобладающий тип сделок в октябре 2021: {dominant_deal_type}')


def question5():
    may_deals = data[
        (data['receiving_date'].dt.month == 6) & (data['receiving_date'].dt.year == 2021) & (
                data['new/current'] == 'Новая')]
    num_originals_received_june = may_deals.shape[0]
    print(f"Количество оригиналов договоров по майским сделкам, полученных в июне 2021: {num_originals_received_june}")


def calculate_bonus(row):
    if row['new/current'] == 'новая':
        if row['status'] == 'ОПЛАЧЕНО' and pandas.notna(row['receiving_date']) and row['receiving_date'] <= reference_date:
            return row['sum'] * 0.07
        else:
            return 0
    elif row['new/current'] == 'текущая':
        if row['status'] != 'ПРОСРОЧЕНО' and pandas.notna(row['receiving_date']) and row['receiving_date'] <= reference_date:
            if row['sum'] > 10000:
                return row['sum'] * 0.05
            else:
                return row['sum'] * 0.03
        else:
            return 0
    else:
        return 0


def task():
    data['bonus'] = data.apply(calculate_bonus, axis=1)
    bonus_summary = data.groupby('sale')['bonus'].sum().reset_index()
    print("Остаток бонусов менеджеров на 01.07.2021:")
    print(bonus_summary)


if __name__ == '__main__':
    question1()
    question2()
    question3()
    question4()
    question5()
    task()
