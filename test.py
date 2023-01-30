def has_invalid_casing(words) -> bool:
    is_all_caps = True
    for word in words:
        for i in range(len(word)):
            if i == 0 and word[i].islower():
                return True
            elif i != 0 and word[i].islower():
                is_all_caps = False
    return True if is_all_caps else False

print(has_invalid_casing(["52F11", "BREAKER", "Status"]))