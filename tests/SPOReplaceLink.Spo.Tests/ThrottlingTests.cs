using System.Net;
using Microsoft.Extensions.Time.Testing;
using SPOReplaceLink.Core.Processing;
using SPOReplaceLink.Spo.Throttling;

namespace SPOReplaceLink.Spo.Tests;

public class ThrottlingTests
{
    [Fact]
    public async Task BudgetWaitsForRefill()
    {
        var time = new FakeTimeProvider(DateTimeOffset.UtcNow);
        var budget = new RuBudget("a", 60, TimeSpan.FromMinutes(1), 0, time);
        await budget.WaitAsync(60, CancellationToken.None);

        var pending = budget.WaitAsync(30, CancellationToken.None).AsTask();
        time.Advance(TimeSpan.FromSeconds(10));
        await Task.Delay(50);
        Assert.False(pending.IsCompleted);
        time.Advance(TimeSpan.FromSeconds(25));
        await pending.WaitAsync(TimeSpan.FromSeconds(5));
        Assert.Equal(90, budget.SpentTotal);
    }

    [Fact]
    public async Task ScaleLowersTheLimit()
    {
        var time = new FakeTimeProvider(DateTimeOffset.UtcNow);
        var budget = new RuBudget("a", 100, TimeSpan.FromMinutes(1), 0, time) { Scale = 0.8 };
        Assert.Equal(80, budget.Capacity);
        await budget.WaitAsync(80, CancellationToken.None);
        var pending = budget.WaitAsync(1, CancellationToken.None).AsTask();
        await Task.Delay(50);
        Assert.False(pending.IsCompleted);
        time.Advance(TimeSpan.FromSeconds(1));
        await pending.WaitAsync(TimeSpan.FromSeconds(5));
    }

    [Fact]
    public async Task WindowLoadShowsSpendingBelowTheLimit()
    {
        var time = new FakeTimeProvider(new DateTimeOffset(2026, 10, 23, 10, 0, 0, TimeSpan.Zero));
        var budget = new RuBudget("a", 1000, TimeSpan.FromMinutes(1), 0, time);
        await budget.WaitAsync(100, CancellationToken.None);
        time.Advance(TimeSpan.FromSeconds(30));
        await budget.WaitAsync(100, CancellationToken.None);
        // The bucket is refilled, yet 200 of 1000 RU went out in the last minute.
        Assert.Equal(0, budget.Utilization, 2);
        Assert.Equal(0.2, budget.WindowLoad, 2);

        time.Advance(TimeSpan.FromSeconds(45));
        Assert.Equal(0.125, budget.WindowLoad, 3);
        time.Advance(TimeSpan.FromMinutes(1));
        Assert.Equal(0, budget.WindowLoad, 3);
    }

    [Fact]
    public async Task DailyLimitPausesUntilTheWindowMoves()
    {
        var time = new FakeTimeProvider(new DateTimeOffset(2026, 10, 23, 10, 0, 0, TimeSpan.Zero));
        var budget = new RuBudget("a", 1000, TimeSpan.FromMinutes(1), 100, time);
        await budget.WaitAsync(100, CancellationToken.None);
        Assert.Equal(100, budget.SpentLast24h);

        var pending = budget.WaitAsync(10, CancellationToken.None).AsTask();
        time.Advance(TimeSpan.FromHours(1));
        await Task.Delay(50);
        Assert.False(pending.IsCompleted);
        time.Advance(TimeSpan.FromHours(23) + TimeSpan.FromMinutes(2));
        await pending.WaitAsync(TimeSpan.FromSeconds(5));
        Assert.Equal(10, budget.SpentLast24h);
    }

    [Fact]
    public async Task ThrottlePausesAllRequestsAndHalvesConcurrency()
    {
        var time = new FakeTimeProvider(DateTimeOffset.UtcNow);
        var state = new ThrottleState("a", 8, TimeSpan.FromMinutes(15), time);
        state.OnThrottled(TimeSpan.FromSeconds(30), is503: false);
        Assert.Equal(4, state.ConcurrencyLimit);
        Assert.True(state.IsPaused);

        var waiting = state.EnterAsync(CancellationToken.None).AsTask();
        await Task.Delay(50);
        Assert.False(waiting.IsCompleted);
        time.Advance(TimeSpan.FromSeconds(31));
        await waiting.WaitAsync(TimeSpan.FromSeconds(5));
        state.Exit();

        for (int i = 0; i < 20; i++)
            state.OnSuccess();
        Assert.Equal(5, state.ConcurrencyLimit);
    }

    [Fact]
    public async Task LimitBlocksExtraRequestsUntilOneLeaves()
    {
        var state = new ThrottleState("a", 1, TimeSpan.FromMinutes(15));
        await state.EnterAsync(CancellationToken.None);
        var second = state.EnterAsync(CancellationToken.None).AsTask();
        await Task.Delay(50);
        Assert.False(second.IsCompleted);
        state.Exit();
        await second.WaitAsync(TimeSpan.FromSeconds(5));
    }

    [Fact]
    public void Only503ForLongMeansBlocked()
    {
        var time = new FakeTimeProvider(DateTimeOffset.UtcNow);
        var state = new ThrottleState("a", 4, TimeSpan.FromMinutes(10), time);
        state.OnThrottled(TimeSpan.FromSeconds(1), is503: true);
        time.Advance(TimeSpan.FromMinutes(9));
        state.OnThrottled(TimeSpan.FromSeconds(1), is503: true);
        Assert.False(state.IsBlocked);
        time.Advance(TimeSpan.FromMinutes(2));
        Assert.True(state.IsBlocked);

        state.OnSuccess();
        Assert.False(state.IsBlocked);
        state.OnThrottled(TimeSpan.FromSeconds(1), is503: false);
        time.Advance(TimeSpan.FromMinutes(20));
        Assert.False(state.IsBlocked);
    }

    [Theory]
    [InlineData(401, null, ErrorClass.Auth)]
    [InlineData(403, null, ErrorClass.Auth)]
    [InlineData(412, null, ErrorClass.Conflict)]
    [InlineData(409, "resourceLocked", ErrorClass.Locked)]
    [InlineData(423, "resourceLocked", ErrorClass.Locked)]
    [InlineData(429, null, ErrorClass.Transient)]
    [InlineData(503, null, ErrorClass.Transient)]
    [InlineData(502, null, ErrorClass.Transient)]
    [InlineData(400, "invalidRequest", ErrorClass.Internal)]
    [InlineData(400, "notAllowed_checkedOut", ErrorClass.Locked)]
    public void ClassifiesErrors(int status, string? code, ErrorClass expected) =>
        Assert.Equal(expected, SpoException.Classify((HttpStatusCode)status, code));
}
