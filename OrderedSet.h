#pragma once

#include <vector>
#include <unordered_map>

template<typename Key>
class OrderedSet {
private:
    std::vector<Key> keys;
    std::unordered_map<Key, size_t> map;

public:
    OrderedSet(std::initializer_list<Key> initList) {
        for (const auto& it : initList) {
            add(it);
        }
    }

    void add(const Key& key) {
        keys.push_back(key);
        map[key] = keys.size() - 1;
    }

    bool contains(const Key& key) const {
        return map.find(key) != map.end();
    }

    typename std::unordered_map<Key, size_t>::const_iterator getByKey(const Key& key) const {
        return map.find(key);
    }

    const Key& getKeyByIndex(const size_t index) const {
        return keys[index];
    }

    typename std::unordered_map<Key, size_t>::const_iterator end() const noexcept {
        return map.end();
    }
};

