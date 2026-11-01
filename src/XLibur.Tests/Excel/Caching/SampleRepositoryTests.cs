using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel.Caching;
using XLibur.Extensions;

namespace XLibur.Tests.Excel.Caching;

public class BaseRepositoryTests
{
    [Test]
    public async Task DifferentEntitiesWithSameKeyStoredOnce()
    {
        // Arrange
        var key = 12345;
        var entity1 = new SampleEntity(key);
        var entity2 = new SampleEntity(key);
        var sampleRepository = CreateSampleRepository();

        // Act
        var storedEntity1 = sampleRepository.Store(ref key, entity1);
        var storedEntity2 = sampleRepository.Store(ref key, entity2);

        // Assert
        await Assert.That(storedEntity1).IsSameReferenceAs(entity1);
        await Assert.That(storedEntity2).IsSameReferenceAs(entity1);
        await Assert.That(storedEntity2).IsNotSameReferenceAs(entity2);
    }

    [Test]
    public async Task ConcurrentAddingCausesNoDuplication()
    {
        // Arrange
        const int countUnique = 30;
        const int repeatCount = 1000;
        var entities = new SampleEntity[countUnique * repeatCount];
        for (var i = 0; i < countUnique; i++)
        {
            for (var j = 0; j < repeatCount; j++)
            {
                entities[i * repeatCount + j] = new SampleEntity(i);
            }
        }

        var sampleRepository = CreateSampleRepository();

        // Act
        Parallel.ForEach(entities, new ParallelOptions { MaxDegreeOfParallelism = 8 },
            e =>
            {
                var key = e.Key;
                sampleRepository.Store(ref key, e);
            });
        var storedEntries = sampleRepository.ToList();

        // Assert
        await Assert.That(storedEntries.Count).IsEqualTo(countUnique);
        await Assert.That(entities).IsNotNull(); // To protect them from GC
    }

    [Test]
    public async Task ReplaceKeyInRepository()
    {
        // Arrange
        var key1 = 12345;
        var key2 = 54321;
        var entity = new SampleEntity(key1);
        var sampleRepository = CreateSampleRepository();
        var storedEntity1 = sampleRepository.Store(ref key1, entity);

        // Act
        sampleRepository.Replace(ref key1, ref key2);
        var containsOld = sampleRepository.ContainsKey(ref key1, out _);
        var containsNew = sampleRepository.ContainsKey(ref key2, out _);
        var storedEntity2 = sampleRepository.GetOrCreate(ref key2);

        // Assert
        await Assert.That(containsOld).IsFalse();
        await Assert.That(containsNew).IsTrue();
        await Assert.That(storedEntity1).IsSameReferenceAs(entity);
        await Assert.That(storedEntity2).IsSameReferenceAs(entity);
    }

    [Test]
    public async Task ConcurrentReplaceKeyInRepository()
    {
        var sampleRepository = new EditableRepository();
        var keys = Enumerable.Range(0, 1000).ToArray();

        // Held, not discarded. The repository stores values through WeakReference, so it only
        // promises that a key keeps mapping to the same instance while something else still
        // reaches that instance. Dropping the seeded entities here left them collectible, and a
        // GC between Replace and the GetOrCreate below then legitimately produced a different
        // instance - failing an assertion the API never made. That made this test fail roughly
        // one full-suite run in four on a branch that only changed allocation elsewhere, and pass
        // 15/15 in isolation, which is what a GC-timing dependency looks like.
        // ConcurrentAddingCausesNoDuplication above already keeps its entities alive for the same
        // reason; this test was the one that did not.
        var seeded = new List<EditableEntity>(keys.Length);
        keys.ForEach(key => seeded.Add(sampleRepository.GetOrCreate(ref key)));

        // Parallel.ForEach takes an Action, and `ref` arguments are not allowed in an async
        // lambda, so mismatches are collected here and asserted once the loop completes.
        var mismatchedKeys = new System.Collections.Concurrent.ConcurrentBag<int>();
        Parallel.ForEach(keys, key =>
        {
            var modifiedKey = key + 2000;
            var val1 = sampleRepository.Replace(ref key, ref modifiedKey);
            var val2 = sampleRepository.GetOrCreate(ref modifiedKey);
            if (!ReferenceEquals(val2, val1))
            {
                mismatchedKeys.Add(key);
            }
        });

        await Assert.That(mismatchedKeys).IsEmpty();

        // The seeded entities must outlive the assertion, or the JIT is free to collect them
        // partway through and reintroduce exactly the race this list exists to remove.
        GC.KeepAlive(seeded);
    }

    [Test]
    public async Task ReplaceNonExistingKeyInRepository()
    {
        var key1 = 100;
        var key2 = 200;
        var key3 = 300;
        var entity = new SampleEntity(key1);
        var sampleRepository = CreateSampleRepository();
        sampleRepository.Store(ref key1, entity);

        sampleRepository.Replace(ref key2, ref key3);
        var all = sampleRepository.ToList();

        await Assert.That(all.Count).IsEqualTo(1);
        await Assert.That(all.First()).IsSameReferenceAs(entity);
    }

    private static SampleRepository CreateSampleRepository()
    {
        return new SampleRepository();
    }

    /// <summary>
    /// Class under testing
    /// </summary>
    private class SampleRepository : XLRepositoryBase<int, SampleEntity>
    {
        public SampleRepository() : base(key => new SampleEntity(key))
        {
        }
    }

    public class SampleEntity
    {
        public int Key { get; private set; }

        public SampleEntity(int key)
        {
            Key = key;
        }
    }

    /// <summary>
    /// Class under testing
    /// </summary>
    private class EditableRepository : XLRepositoryBase<int, EditableEntity>
    {
        public EditableRepository() : base(_ => new EditableEntity())
        {
        }
    }

    private class EditableEntity;
}
