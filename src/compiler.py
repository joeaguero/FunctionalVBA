def name_of_abstract_level(level: int) -> str:
    return f"I{name_of_level(level)}"

def name_of_level(level: int) -> str:
    return f"Lambda{level}" if level > 0 else "Nullary"

def build_header(parent_level: int) -> str:
    with open("templates/concrete/HeaderPart.cls", 'r') as f:
        return f.read().format(name=name_of_level(parent_level))

def build_inheritances(sublevels: list[int]) -> str:
    return "\r".join(map(lambda a: f"Implements {name_of_abstract_level(a)}", sublevels))

def build_data(arguments: list[str]) -> str:
    with open("templates/concrete/DataPart.cls",'r') as f:
        return (
            f.read()
             .format(members="".join(
                map(lambda a: f"\r  {a} As Variant", arguments))))

def build_constructors(arguments: list[int]) -> str:
    with open("templates/concrete/ConstructorPart.cls",'r') as f:
        template: str = f.read()

    def build_constructor(level: int) -> str:
        filtered: list[str] = arguments[:len(arguments) - level]
        return template.format(
            level=level,
            abstract_name=name_of_abstract_level(level),
            members="".join(map(lambda a: f", ByRef {a} As Variant", filtered)),
            members_assignment="".join(map(lambda a: f"\r  LetSet(This.{a}) = {a}", filtered)))

    return "\r\r".join(map(build_constructor,range(len(arguments), -1, -1)))

def build_runners(arguments: list[int]) -> str:
    with open("templates/concrete/RunPart.cls",'r') as f:
        lambda_template: str = f.read()

    with open("templates/concrete/NullaryPart.cls", 'r') as f:
        nullary_template: str = f.read()

    def build_runner(level: int) -> str:
        filtered: list[str] = arguments[:len(arguments) - level]
        next_level: int = level - 1

        if level == 0:
            return nullary_template.format(
                members="".join(map(lambda a: f", This.{a}", filtered)))

        return lambda_template.format(
            abstract_name=name_of_abstract_level(level),
            next_name=name_of_abstract_level(next_level),
            parent_name=name_of_level(len(arguments)),
            next_level=next_level,
            members="".join(map(lambda a: f", This.{a}", filtered)))

    return "\r".join(map(build_runner, range(len(arguments), -1, -1)))


def build_terminate(arguments: list[int]) -> str:
    with open("templates/concrete/CleanupPart.cls", 'r') as f:
        return (
            f.read()
             .format(members="".join(
                map(lambda a: f"\r  This.{a} = Empty", arguments))))

def build_concrete(level: int) -> None:
    sublevels: list[int] = list(range(level, -1, -1))
    arguments: list[str] = list(map(lambda a: f"Arg{a}", filter(lambda a: a > 0, reversed(sublevels))))

    sections: list[str] = [
        build_header(level),
        build_inheritances(sublevels),
        build_data(arguments),
        build_constructors(arguments),
        build_runners(arguments) ]

    if arguments:
        sections.append(build_terminate(arguments))

    return "\r\r".join(sections)

def build_class_header() -> str:
    with open("templates/class/HeaderPart.bas",'r') as f:
        return f.read()

def build_class_constructors(levels: int) -> str:
    with open("templates/class/ConstructorPart.bas", 'r') as f:
        template: str = f.read()

    def build_class_constructor(level: int) -> str:
        return template.format(
            name=name_of_level(level),
            abstract_name=name_of_abstract_level(level),
            level=level)

    return "\r".join(map(build_class_constructor, range(levels + 1)))

def build_class(levels: int) -> str:
    sections: list[str] = [
        build_class_header(),
        build_class_constructors(levels)
    ]

    return "\r".join(sections)

def build_abstraction(level: int) -> str:
    if level == 0:
        with open("templates/abstraction/INullary.cls", 'r') as f:
            return f.read()

    with open("templates/abstraction/ILambda.cls", 'r') as f:
        return f.read().format(
            current_name=name_of_abstract_level(level),
            next_name=name_of_abstract_level(level - 1))

def main() -> None:
    levels: int = 6

    for level in range(levels + 1):
        with open(f"../lambda/{name_of_abstract_level(level)}.cls", 'w') as f:
            f.write(build_abstraction(level))

        with open(f"../lambda/{name_of_level(level)}.cls", 'w') as f:
            f.write(build_concrete(level))

    with open("../lambda/FunctionalConstructors.bas", 'w') as f:
        f.write(build_class(levels))

if __name__ == "__main__":
    main()
