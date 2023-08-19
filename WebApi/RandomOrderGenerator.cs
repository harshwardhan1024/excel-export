using Bogus;
using WebApi.Models;

namespace WebApi;

public class RandomOrderGenerator
{
    public List<Order> Generate(int count)
    {
        var idCounter = 1;
        return new Faker<Order>()
            .RuleFor(o => o.Id, f => idCounter++)
            .RuleFor(o => o.Amount, f => f.Random.Decimal(10m, 100m))
            .RuleFor(o => o.HasShipped, f => f.Random.Bool())
            .RuleFor(o => o.Date, (f, o) => f.Date.Past())
            .RuleFor(o => o.UserEmail, f => f.Internet.Email())
            .Generate(count);
    }
}
