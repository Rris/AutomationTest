
public class NYStylePizzaStore extends PizzaStore{
	@Override
	Pizza createPizza(String type) {
		Pizza pizza = null;
		if (type.equals("cheese")) {
			pizza = new NYStyleCheesePizza();
		}
		return pizza;
	}

}
