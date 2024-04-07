using ClosedXML.Excel;
using Akelon_case_3.Models;

class Program
{
    static void Main()
    {
        Console.WriteLine("Введите путь до файла с данными: ");
        string? inputPath = Console.ReadLine();

        string? path = inputPath?.Replace("\"", "");

        using (var workbook = new XLWorkbook(path))
        {
            var productsWorksheet = workbook.Worksheet(1);
            var clientsWorksheet = workbook.Worksheet(2);
            var ordersWorksheet = workbook.Worksheet(3);

            var products = ReadProductsFromWorksheet(productsWorksheet);
            var clients = ReadClientsFromWorksheet(clientsWorksheet);
            var orders = ReadOrdersFromWorksheet(ordersWorksheet);

            while (true)
            {
                Console.WriteLine("Выберите команду:");
                Console.WriteLine("1. Вывести информацию о клиентах, заказавших товар");
                Console.WriteLine("2. Изменить контактное лицо клиента");
                Console.WriteLine("3. Определить золотого клиента с наибольшим количеством заказов за указанный год, месяц");
                Console.WriteLine("4. Выйти из программы");

                int choice;
                if (!int.TryParse(Console.ReadLine(), out choice) || choice < 1 || choice > 4)
                {
                    Console.WriteLine("Некорректный ввод. Пожалуйста, введите число от 1 до 4.");
                    continue;
                }

                switch (choice)
                {
                    case 1:
                        PrintClientsByProduct(products, orders, clients);
                        break;
                    case 2:
                        ChangeContactPerson(clients, path);
                        break;
                    case 3:
                        PrintGoldenClientAndMostOrdersClient(orders, clients);
                        break;
                    case 4:
                        return;
                }
            }
        }
    }

    static List<Product> ReadProductsFromWorksheet(IXLWorksheet worksheet)
    {
        var products = new List<Product>();
        foreach (var row in worksheet.RowsUsed().Skip(1))
        {
            products.Add(new Product
            {
                Code = row.Cell(1).GetValue<int>(),
                Name = row.Cell(2).GetString(),
                Unit = row.Cell(3).GetString(),
                Price = row.Cell(4).GetValue<decimal>()
            });
        }

        return products;
    }

    static List<Client> ReadClientsFromWorksheet(IXLWorksheet worksheet)
    {
        var clients = new List<Client>();
        foreach (var row in worksheet.RowsUsed().Skip(1))
        {
            clients.Add(new Client
            {
                Code = row.Cell(1).GetValue<int>(),
                OrganizationName = row.Cell(2).GetString(),
                Address = row.Cell(3).GetString(),
                ContactPerson = row.Cell(4).GetString()
            });
        }

        return clients;
    }

    static List<Order> ReadOrdersFromWorksheet(IXLWorksheet worksheet)
    {
        var orders = new List<Order>();
        foreach (var row in worksheet.RowsUsed().Skip(1))
        {
            orders.Add(new Order
            {
                OrderCode = row.Cell(1).GetValue<int>(),
                ProductCode = row.Cell(2).GetValue<int>(),
                ClientCode = row.Cell(3).GetValue<int>(),
                OrderNumber = row.Cell(4).GetValue<int>(),
                Quantity = row.Cell(5).GetValue<int>(),
                Date = row.Cell(6).GetDateTime()
            });
        }

        return orders;
    }


    /// <summary>
    /// Получение информации о заказе
    /// </summary>
    /// <param name="products">Товары</param>
    /// <param name="orders">Заявки</param>
    static void PrintClientsByProduct(List<Product> products, List<Order> orders, List<Client> clients)
    {
        Console.WriteLine("Введите наименование товара:");
        string? productName = Console.ReadLine();

        var product = products.Select(p => p)
                              .FirstOrDefault(p => p.Name.Equals(productName, StringComparison.OrdinalIgnoreCase));

        if (product == null)
        {
            Console.WriteLine($"Товар с наименованием '{productName}' не найден.");
            return;
        }

        var relevantOrders = orders.Where(o => o.ProductCode == product.Code)
                                   .Select(o => o)
                                   .ToList();

        var contactPersonCodes = relevantOrders.Select(c => c.ClientCode)
                                               .ToList();

        var contactPerson = clients.Where(c => contactPersonCodes.Contains(c.Code))
                                   .Select(c => c.ContactPerson)
                                   .FirstOrDefault();

        if (!relevantOrders.Any())
        {
            Console.WriteLine($"Нет заказов для товара '{productName}'.");
            return;
        }

        Console.WriteLine($"Информация о клиентах, заказавших товар '{productName}':");

        foreach (var order in relevantOrders)
        {
            Console.WriteLine($"Код клиента: {order.ClientCode}, Контактное лицо: {contactPerson}, Количество: {order.Quantity}, Цена: {product.Price}, Дата заказа: {order.Date}");
        }
    }


    /// <summary>
    /// Изменение контактного лица клиента по названию организации
    /// </summary>
    /// <param name="clients">Клиенты</param>
    static void ChangeContactPerson(List<Client> clients, string filePath)
    {
        Console.WriteLine("Введите наименование организации клиента:");
        string? organizationName = Console.ReadLine();
        Console.WriteLine("Введите ФИО нового контактного лица:");
        string? newContactPerson = Console.ReadLine();

        var client = clients.FirstOrDefault(c => c.OrganizationName.Equals(organizationName, StringComparison.OrdinalIgnoreCase));
        if (client == null)
        {
            Console.WriteLine($"Клиент с наименованием организации '{organizationName}' не найден.");
            return;
        }

        client.ContactPerson = newContactPerson;

        using (var workbook = new XLWorkbook(filePath))
        {
            var clientsWorksheet = workbook.Worksheet(2);

            var row = clientsWorksheet.RowsUsed().FirstOrDefault(r => r.Cell(2).GetString().Equals(organizationName, StringComparison.OrdinalIgnoreCase));

            if (row != null)
            {
                row.Cell(4).Value = newContactPerson;
                workbook.Save();
            }
            else
            {
                Console.WriteLine($"Строка с наименованием организации '{organizationName}' не найдена в файле.");
            }
        }

        Console.WriteLine($"Контактное лицо клиента '{organizationName}' изменено на '{newContactPerson}'.");
    }



    /// <summary>
    /// Определение золотого клиента
    /// </summary>
    /// <param name="orders">Заявки</param>
    static void PrintGoldenClientAndMostOrdersClient(List<Order> orders, List<Client> clients)
    {
        Console.WriteLine("Введите год:");
        int year = int.Parse(Console.ReadLine());
        Console.WriteLine("Введите месяц:");
        int month = int.Parse(Console.ReadLine());

        var relevantOrders = orders.Where(o => o.Date.Year == year &&
                                               o.Date.Month == month)
                                   .Select(o => o)
                                   .ToList();

        if (relevantOrders.Count == 0)
        {
            Console.WriteLine($"Записей с такими параметрами не найдено.");
            return;
        }

        var mostOrdersClient = relevantOrders.GroupBy(o => o.ClientCode)
                                             .OrderByDescending(g => g.Count())
                                             .FirstOrDefault()?.Key;

        var contactPerson = clients.Where(c => c.Code == mostOrdersClient)
                                   .Select(c => c.ContactPerson)
                                   .FirstOrDefault();

        if (mostOrdersClient > 0)
        {
            Console.WriteLine($"Золотой клиент за {month}/{year}: Код клиента: {mostOrdersClient}, Контактное лицо: {contactPerson}");
        }
    }
}